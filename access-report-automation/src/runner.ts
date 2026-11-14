import { createHash } from 'node:crypto';
import { gunzipSync, gzipSync } from 'node:zlib';
import { CertificateCredential, Credential, FederatedCredential, TokenProvider } from './auth.js';
import { Config, ConfigError, Env, parsePrincipalsCsv, PRINCIPALS_TEMPLATE, siteFilter } from './config.js';
import { Graph, PrincipalRequest, SiteRef, UnresolvedPrincipal } from './graph.js';
import { Http } from './http.js';
import { Logger } from './log.js';
import { Principal } from './permissions.js';
import { buildReport, FailedSite, Snapshot, unpackSnapshot } from './report.js';
import { AccessRow, isSiteProblem, scanError, SiteScanner } from './scanner.js';
import { SharePoint, trimSlash } from './sharepoint.js';
import { LocalStore, SharePointStore, Store } from './store.js';

/**
 * One run: work out who to report on, read every site, write the report.
 *
 * Runs are restartable. Progress is saved next to the reports every few minutes; a run that
 * hits its time limit, loses its connection or is cancelled is picked up by the next one.
 * A month that already has its report is left alone, so the job can safely be scheduled
 * several times a month (which is what makes it finish on large tenants, and retry by itself).
 */

export interface Clients {
  http: Http;
  tokens: TokenProvider;
  credential: Credential;
  graph: Graph;
  sp: SharePoint;
}

export function createClients(cfg: Config, log: Logger, env: Env, httpOverrides: { baseDelayMs?: number; maxRetries?: number } = {}): Clients {
  const http = new Http({ concurrency: cfg.scan.concurrency, log, ...httpOverrides });
  const credential: Credential =
    cfg.auth.mode === 'certificate'
      ? new CertificateCredential(cfg.clientId, String(cfg.auth.certificate), cfg.auth.password)
      : new FederatedCredential(http, { requestUrl: env.ACTIONS_ID_TOKEN_REQUEST_URL, requestToken: env.ACTIONS_ID_TOKEN_REQUEST_TOKEN, tokenFile: env.AZURE_FEDERATED_TOKEN_FILE });
  const tokens = new TokenProvider(http, cfg.endpoints.login, cfg.tenantId, cfg.clientId, credential);
  return { http, tokens, credential, graph: new Graph(http, tokens, cfg.endpoints.graph), sp: new SharePoint(http, tokens) };
}

export interface RunOptions {
  /** "2026-11". Default: the current month (UTC). */
  period?: string;
  /** Make the report again even when this month's exists. */
  force?: boolean;
  /** false: write everything to `outDir` instead of SharePoint (trial runs). */
  upload?: boolean;
  outDir?: string;
  /** Only the first N sites (trial runs). */
  limit?: number;
  now?: () => number;
}

export interface RunResult {
  status: 'done' | 'skipped' | 'partial';
  period: string;
  sitesTotal: number;
  sitesDone: number;
  sitesFailed: number;
  principals: number;
  unresolved: number;
  rows: number;
  added: number;
  removed: number;
  changed: number;
  comparedWith?: string;
  files: string[];
  warnings: string[];
  requests: number;
  throttled: number;
}

interface Checkpoint {
  v: 1;
  period: string;
  fingerprint: string;
  sites: Array<[string, string]>;
  done: number[];
  failed: Array<{ i: number; error: string; site: boolean }>;
  rows: AccessRow[];
  /** Notes for the reader of the report that came up while scanning. */
  notes?: string[];
  startedAt: string;
}

export function currentPeriod(nowMs: number): string {
  return new Date(nowMs).toISOString().substring(0, 7);
}

const STATE = '_state';
const snapshotName = (period: string): string => `${STATE}/snapshot-${period}.json.gz`;
const checkpointName = (period: string): string => `${STATE}/checkpoint-${period}.json.gz`;

/** A change to who or what is scanned makes saved progress useless: it is thrown away, not mixed. */
export function fingerprint(cfg: Config, principals: Principal[], limit: number | undefined): string {
  const basis = {
    principals: principals.map((p) => p.key).sort(),
    scan: [cfg.scan.includeLibraries, cfg.scan.includeLists, cfg.scan.includeSubsites, cfg.scan.expandGroups, cfg.scan.thorough],
    sites: [cfg.sites.only, cfg.sites.include, cfg.sites.exclude, cfg.sites.includeOneDrive],
    limit: limit || 0
  };
  return createHash('sha256').update(JSON.stringify(basis)).digest('hex').substring(0, 16);
}

function scopeText(cfg: Config): string {
  const parts = ['sites'];
  if (cfg.scan.includeSubsites) {
    parts.push('subsites');
  }
  if (cfg.scan.includeLibraries) {
    parts.push('libraries');
  }
  if (cfg.scan.includeLists) {
    parts.push('lists');
  }
  let s = parts.join(', ');
  if (cfg.sites.only.length) {
    s += `; only ${cfg.sites.only.length} listed site${cfg.sites.only.length === 1 ? '' : 's'}`;
  } else {
    s += cfg.sites.includeOneDrive ? '; all site collections including OneDrive' : '; all site collections (OneDrive left out)';
    if (cfg.sites.include.length || cfg.sites.exclude.length) {
      s += ', filtered';
    }
  }
  return s;
}

async function loadPrincipalRequests(cfg: Config, store: Store, log: Logger): Promise<PrincipalRequest[]> {
  const requests = cfg.principals.slice();
  if (cfg.principalsFile) {
    const data = await store.read(cfg.principalsFile);
    if (data) {
      const parsed = parsePrincipalsCsv(Buffer.from(data).toString('utf8'));
      parsed.problems.forEach((p) => log.warn(p));
      parsed.principals.forEach((p) => requests.push(p));
      log.info(`${cfg.principalsFile}: ${parsed.principals.length} entr${parsed.principals.length === 1 ? 'y' : 'ies'}.`);
    } else if (!requests.length) {
      await store.write(cfg.principalsFile, Buffer.from('﻿' + PRINCIPALS_TEMPLATE, 'utf8'));
      throw new ConfigError(
        `Nobody to report on yet. An empty ${cfg.principalsFile} was created in the report folder (${store.description}): add one row per person or group (Type = user or group; Identity = e-mail / sign-in name, or a group's name or object id), then run again.`
      );
    }
  }
  if (!requests.length) {
    throw new ConfigError(`Nobody to report on. Add people or groups to ${cfg.principalsFile || 'the "principals" setting'}.`);
  }
  const seen = new Set<string>();
  return requests.filter((r) => {
    const k = `${r.type}:${r.id.toLowerCase()}`;
    if (seen.has(k)) {
      return false;
    }
    seen.add(k);
    return true;
  });
}

async function previousSnapshot(store: Store, period: string, log: Logger): Promise<Snapshot | undefined> {
  const names = (await store.list(STATE))
    .map((n) => /^snapshot-(\d{4}-\d{2})\.json\.gz$/.exec(n))
    .filter((m): m is RegExpExecArray => !!m && m[1] < period)
    .map((m) => m[1])
    .sort();
  // Newest first; a damaged file falls back to the one before it.
  for (let i = names.length - 1; i >= 0; i--) {
    try {
      const data = await store.read(snapshotName(names[i]));
      if (data) {
        return unpackSnapshot(data);
      }
    } catch (e) {
      log.warn(`The saved ${names[i]} report could not be read for comparison (${(e as Error).message}).`);
    }
  }
  return undefined;
}

export async function run(cfg: Config, clients: Clients, log: Logger, opts: RunOptions = {}): Promise<RunResult> {
  const now = opts.now ?? Date.now;
  const started = now();
  const period = opts.period ?? currentPeriod(started);
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(period)) {
    throw new ConfigError(`The period must look like 2026-11, not "${period}".`);
  }
  const upload = opts.upload !== false;
  const store: Store = upload ? new SharePointStore(clients.sp, cfg.report.siteUrl, cfg.report.libraryPath, cfg.report.folder) : new LocalStore(opts.outDir || 'out');
  if (store instanceof SharePointStore) {
    await store.prepare();
  }
  log.info(`Report folder: ${store.description}`);

  const result: RunResult = { status: 'done', period, sitesTotal: 0, sitesDone: 0, sitesFailed: 0, principals: 0, unresolved: 0, rows: 0, added: 0, removed: 0, changed: 0, files: [], warnings: [], requests: 0, throttled: 0 };
  const finish = (): RunResult => {
    result.requests = clients.http.stats.requests;
    result.throttled = clients.http.stats.throttled;
    return result;
  };

  if (!opts.force && (await store.list(STATE)).indexOf(snapshotName(period).split('/')[1]) >= 0) {
    log.info(`The ${period} report already exists. Nothing to do (use --force to make it again).`);
    result.status = 'skipped';
    return finish();
  }

  // ---- who ----
  const requests = await loadPrincipalRequests(cfg, store, log);
  const resolved = await clients.graph.resolvePrincipals(requests, cfg.scan.expandGroups);
  const principals = resolved.principals;
  const unresolved: UnresolvedPrincipal[] = resolved.unresolved;
  const warnings = resolved.warnings.slice();
  unresolved.forEach((u) => log.warn(`${u.type} "${u.input}": ${u.reason}`));
  if (!principals.length) {
    throw new ConfigError('None of the people or groups in the list were found in the directory.');
  }
  const groupsExpanded = cfg.scan.expandGroups && !resolved.warnings.length;
  log.info(`${principals.length} ${principals.length === 1 ? 'person or group' : 'people and groups'} to report on${unresolved.length ? `, ${unresolved.length} not found` : ''}.`);
  if (clients.credential instanceof CertificateCredential) {
    const days = clients.credential.daysLeft(started);
    if (days < 45) {
      const w = `The sign-in certificate expires in ${days} day${days === 1 ? '' : 's'} (${clients.credential.expires.toISOString().substring(0, 10)}). Replace it before then, or the monthly report stops.`;
      warnings.push(w);
      log.warn(w);
    }
  }
  result.principals = principals.length;
  result.unresolved = unresolved.length;

  // ---- where ----
  const print = fingerprint(cfg, principals, opts.limit);
  let cp: Checkpoint | undefined;
  if (!opts.force) {
    try {
      const data = await store.read(checkpointName(period));
      if (data) {
        const saved = JSON.parse(gunzipSync(data).toString('utf8')) as Checkpoint;
        if (saved && saved.v === 1 && saved.period === period && saved.fingerprint === print && Array.isArray(saved.sites)) {
          cp = saved;
          log.info(`Resuming: ${saved.done.length} of ${saved.sites.length} sites were already read.`);
        } else {
          log.info('Saved progress belongs to different settings; starting again.');
        }
      }
    } catch (e) {
      log.warn(`Saved progress could not be read (${(e as Error).message}); starting again.`);
    }
  }
  if (!cp) {
    let sites: SiteRef[];
    if (cfg.sites.only.length) {
      sites = cfg.sites.only.map((u) => ({ url: trimSlash(u), title: trimSlash(u) }));
    } else {
      log.info('Reading the list of site collections...');
      try {
        sites = await clients.graph.listSites(cfg.sites.includeOneDrive, (n) => log.debug(`${n} sites listed`));
      } catch (e) {
        throw new Error(`The list of sites could not be read from Microsoft Graph (${(e as Error).message}). The app needs the Graph application permission Sites.Read.All with admin consent; or list the sites to scan in "sites.only".`);
      }
      const keep = siteFilter(cfg.sites.include, cfg.sites.exclude);
      sites = sites.filter((s) => keep(s.url));
    }
    if (opts.limit && opts.limit > 0) {
      sites = sites.slice(0, opts.limit);
    }
    if (!sites.length) {
      throw new ConfigError('There are no sites to scan. Check "sites.include" / "sites.exclude".');
    }
    cp = { v: 1, period, fingerprint: print, sites: sites.map((s) => [s.url, s.title] as [string, string]), done: [], failed: [], rows: [], startedAt: new Date(started).toISOString() };
  }
  const state = cp;
  const total = state.sites.length;
  result.sitesTotal = total;
  log.info(`${total} site collection${total === 1 ? '' : 's'} to check.`);

  // ---- scan ----
  const scanner = new SiteScanner(clients.sp, principals, {
    includeLibraries: cfg.scan.includeLibraries,
    includeLists: cfg.scan.includeLists,
    includeSubsites: cfg.scan.includeSubsites,
    thorough: cfg.scan.thorough,
    groupsExpanded,
    maxSubsites: cfg.scan.maxSubsites
  });
  const earlierNotes = state.notes || [];
  const scanNotes = (): string[] => earlierNotes.concat(Array.from(scanner.warnings.values())).filter((n, i, all) => all.indexOf(n) === i);
  const skipWebs = new Set(state.sites.map((s) => trimSlash(s[0]).toLowerCase()));
  const deadline = cfg.scan.maxMinutes > 0 ? started + cfg.scan.maxMinutes * 60000 : Infinity;
  const done = new Set(state.done);
  let saving: Promise<void> = Promise.resolve();
  let lastSave = now();
  const saveCheckpoint = (): Promise<void> => {
    // One save at a time, each with the state as it is when it starts.
    saving = saving
      .catch(() => undefined)
      .then(async () => {
        state.done = Array.from(done);
        state.notes = scanNotes();
        lastSave = now();
        await store.write(checkpointName(period), gzipSync(Buffer.from(JSON.stringify(state), 'utf8')));
      });
    return saving;
  };

  let stop: string | undefined;
  let outOfTime = false;
  let troubleInARow = 0;
  let lastProgress = now();
  const scanPass = async (queue: number[]): Promise<void> => {
    const worker = async (): Promise<void> => {
      while (!stop && !outOfTime) {
        const i = queue.shift();
        if (i === undefined) {
          return;
        }
        if (now() >= deadline) {
          queue.unshift(i);
          outOfTime = true;
          return;
        }
        const site = { url: state.sites[i][0], title: state.sites[i][1] };
        try {
          const rows = await scanner.scanSite(site, skipWebs);
          rows.forEach((r) => state.rows.push(r));
          done.add(i);
          troubleInARow = 0;
        } catch (e) {
          const siteOnly = isSiteProblem(e);
          const error = scanError(e);
          state.failed.push({ i, error, site: siteOnly });
          done.add(i);
          log.debug(`${site.url}: ${error}`);
          troubleInARow = siteOnly ? 0 : troubleInARow + 1;
          if (troubleInARow >= 8) {
            stop = `Stopped after 8 errors in a row that do not look like single-site problems (last: ${error}). Progress is saved; the next run carries on.`;
          }
        }
        if (now() - lastProgress > 60000) {
          lastProgress = now();
          log.info(`${done.size} of ${total} sites read, ${state.rows.length} places with access so far.`);
        }
        if (now() - lastSave > cfg.scan.checkpointMinutes * 60000) {
          await saveCheckpoint().catch((e) => log.warn(`Progress could not be saved (${(e as Error).message}).`));
        }
      }
    };
    await Promise.all(Array.from({ length: cfg.scan.concurrency }, () => worker()));
  };

  const pending: number[] = [];
  for (let i = 0; i < total; i++) {
    if (!done.has(i)) {
      pending.push(i);
    }
  }
  await scanPass(pending);

  // Sites that failed for a passing reason get one more try at the end.
  const retry = state.failed.filter((f) => !f.site);
  if (!stop && !outOfTime && retry.length) {
    log.info(`Trying ${retry.length} site${retry.length === 1 ? '' : 's'} again.`);
    state.failed = state.failed.filter((f) => f.site);
    retry.forEach((f) => done.delete(f.i));
    troubleInARow = 0;
    await scanPass(retry.map((f) => f.i));
  }

  result.sitesDone = done.size - state.failed.length;
  result.sitesFailed = state.failed.length;
  if (stop || outOfTime) {
    if (stop) {
      // Sites recorded as failed in the run-up to stopping are done again next time.
      state.failed.filter((f) => !f.site).forEach((f) => done.delete(f.i));
      state.failed = state.failed.filter((f) => f.site);
    }
    await saveCheckpoint();
    if (stop) {
      throw new Error(stop);
    }
    log.info(`Time limit reached with ${done.size} of ${total} sites read. Progress is saved; the next run carries on.`);
    result.status = 'partial';
    result.sitesDone = done.size - state.failed.length;
    result.sitesFailed = state.failed.length;
    return finish();
  }
  if (state.failed.length === total) {
    await saveCheckpoint().catch(() => undefined);
    throw new Error(
      `None of the ${total} sites could be read (first problem: ${state.failed[0].error}). No report was written. The app most likely lacks the SharePoint application permission Sites.FullControl.All (with admin consent); run the "check" command for details.`
    );
  }

  // ---- report ----
  const failedSites: FailedSite[] = state.failed.map((f) => ({ url: state.sites[f.i][0], title: state.sites[f.i][1], error: f.error }));
  const previous = await previousSnapshot(store, period, log);
  const files = buildReport({
    name: cfg.report.name,
    period,
    generatedAt: new Date(now()),
    principals: principals.map((p) => ({ key: p.key, label: p.label, kind: p.kind })),
    rows: state.rows,
    failedSites,
    unresolved,
    warnings: (scanNotes().forEach((n) => warnings.push(n)), warnings),
    sitesTotal: total,
    previous,
    scope: scopeText(cfg) + (opts.limit ? `; TRIAL RUN limited to ${opts.limit} sites` : '')
  });
  const base = `${cfg.report.name} ${period}`;
  result.files.push(await store.write(`${base}.xlsx`, files.workbook, { renameIfLocked: true }));
  result.files.push(await store.write(`csv/${base}.csv`, Buffer.from(files.accessCsv, 'utf8'), { renameIfLocked: true }));
  if (files.changesCsv) {
    result.files.push(await store.write(`csv/${base} - changes.csv`, Buffer.from(files.changesCsv, 'utf8'), { renameIfLocked: true }));
  }
  if (files.issuesCsv) {
    result.files.push(await store.write(`csv/${base} - not read.csv`, Buffer.from(files.issuesCsv, 'utf8'), { renameIfLocked: true }));
  }
  // The snapshot goes last: it is what says "this month is done".
  await store.write(snapshotName(period), gzipSync(Buffer.from(JSON.stringify(files.snapshot), 'utf8')));
  await saving.catch(() => undefined);
  await store.remove(checkpointName(period)).catch((e) => log.warn(`Saved progress could not be removed (${(e as Error).message}).`));
  // Progress files of months that never finished are of no use any more.
  for (const n of await store.list(STATE).catch(() => [] as string[])) {
    const m = /^checkpoint-(\d{4}-\d{2})\.json\.gz$/.exec(n);
    if (m && m[1] < period) {
      await store.remove(`${STATE}/${n}`).catch(() => undefined);
    }
  }

  result.rows = files.counts.rows;
  result.added = files.counts.added;
  result.removed = files.counts.removed;
  result.changed = files.counts.changed;
  result.comparedWith = previous ? previous.period : undefined;
  result.warnings = warnings;
  log.info(`Done: ${result.rows} places with access across ${files.counts.sitesWithAccess} sites; ${result.sitesFailed} site${result.sitesFailed === 1 ? '' : 's'} not read.`);
  return finish();
}
