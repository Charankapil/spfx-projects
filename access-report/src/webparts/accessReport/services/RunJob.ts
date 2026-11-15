import { CancelledError } from '../core/SPClient';
import { AccessScanner, ISite } from './AccessScan';
import { ISubject } from './permissions';
import { REPORT_HEADERS, toReportCsv, UNREAD_HEADERS } from './reportCsv';

/**
 * A report run in the browser: every site in scope, for every subject. Progress is saved in
 * the browser so a run stopped by a closed tab, a lost connection or repeated errors carries
 * on from where it was.
 */

export type RunSpeed = 'gentle' | 'normal';

export interface IRunJob {
  v: 2;
  id: string;
  subjects: ISubject[];
  includeLibraries: boolean;
  siteFilter: string;
  folder: string;
  speed: RunSpeed;
  /** Index of the next site to start. */
  next: number;
  /** Sites started but not finished when the job was saved: done again on resume. */
  pending: number[];
  /** [site index, ...report cells] */
  rows: Array<[number, ...string[]]>;
  /** [site index, ...unread cells] */
  unread: Array<[number, ...string[]]>;
  started: number;
  saved: number;
  finished: boolean;
  total: number;
  /** Server-relative URL of the saved report, once it is in the folder. */
  savedAs?: string;
}

const JOB_KEY = 'accessReport.job.v2';
const SITES_KEY = 'accessReport.job.sites.v2';

/** Pause between sites in gentle mode, on top of SPClient's own pacing. */
export const GENTLE_PAUSE_MS = 500;
/** This many sites in a row failing for a reason other than "no access" pauses the run. */
export const MAX_FAILS_IN_A_ROW = 8;
const SAVE_EVERY_MS = 5000;

export function newJob(subjects: ISubject[], includeLibraries: boolean, siteFilter: string, folder: string, speed: RunSpeed, total: number): IRunJob {
  return { v: 2, id: String(Date.now()), subjects, includeLibraries, siteFilter, folder, speed, next: 0, pending: [], rows: [], unread: [], started: Date.now(), saved: Date.now(), finished: false, total };
}

/** The saved job and its site list (kept apart: the list is written once, the job often). */
export function loadJob(): { job: IRunJob; sites: ISite[] } | undefined {
  try {
    const job = JSON.parse(window.localStorage.getItem(JOB_KEY) || 'null') as IRunJob | null;
    const sites = JSON.parse(window.localStorage.getItem(SITES_KEY) || 'null') as Array<[string, string]> | null;
    if (!job || job.v !== 2 || !sites || sites.length !== job.total) {
      return undefined;
    }
    return { job, sites: sites.map(([url, title]) => ({ url, title })) };
  } catch {
    return undefined;
  }
}

/** false when the browser refused (private mode, storage full): the run continues, but cannot resume after a reload. */
export function saveJob(job: IRunJob, sites?: ISite[]): boolean {
  try {
    if (sites) {
      window.localStorage.setItem(SITES_KEY, JSON.stringify(sites.map((s) => [s.url, s.title])));
    }
    job.saved = Date.now();
    window.localStorage.setItem(JOB_KEY, JSON.stringify(job));
    return true;
  } catch {
    return false;
  }
}

export function clearJob(): void {
  try {
    window.localStorage.removeItem(JOB_KEY);
    window.localStorage.removeItem(SITES_KEY);
  } catch {
    /* nothing saved */
  }
}

/** The queue of site indexes still to do: interrupted ones first, then the rest. */
export function remaining(job: IRunJob): number[] {
  const out = job.pending.slice();
  for (let i = job.next; i < job.total; i++) {
    out.push(i);
  }
  return out;
}

export function processed(job: IRunJob): number {
  return Math.max(0, job.next - job.pending.length);
}

/** Site indexes with something unread, each once. */
export function unreadSites(job: IRunJob): number[] {
  const out: number[] = [];
  job.unread.forEach((u) => {
    if (out.indexOf(u[0]) < 0) {
      out.push(u[0]);
    }
  });
  return out;
}

/** Puts the sites that had unread parts back in the queue, dropping what they produced. */
export function requeueUnread(job: IRunJob): void {
  const again = unreadSites(job);
  job.rows = job.rows.filter((r) => again.indexOf(r[0]) < 0);
  job.unread = [];
  job.pending = job.pending.concat(again.filter((i) => job.pending.indexOf(i) < 0));
  job.finished = false;
  job.savedAs = undefined;
}

const bySite = <T extends [number, ...string[]]>(xs: T[]): string[][] =>
  xs
    .map((x, n) => ({ x, n }))
    .sort((a, b) => a.x[0] - b.x[0] || a.n - b.n)
    .map((o) => o.x.slice(1) as string[]);

/** The report file: rows in site order whatever order the sites finished in. */
export function reportCsv(job: IRunJob): string {
  return toReportCsv(REPORT_HEADERS, bySite(job.rows));
}

export function unreadCsv(job: IRunJob): string {
  return toReportCsv(UNREAD_HEADERS, bySite(job.unread));
}

export interface IRunToken {
  stop: boolean;
  /** Set when the run paused itself. */
  reason?: string;
}

/**
 * Runs (or resumes) a job until it is finished, stopped (token.stop) or paused by errors.
 * `onProgress(saved)` is called after every site; `saved` is false when the browser refused
 * to store the progress.
 */
export async function runJob(job: IRunJob, sites: ISite[], scanner: AccessScanner, token: IRunToken, onProgress: (saved: boolean) => void, pauseMs = GENTLE_PAUSE_MS): Promise<void> {
  const queue = remaining(job);
  job.pending = [];
  let failsInARow = 0;
  let lastSave = 0;
  let stored = true;
  const persist = (force: boolean): void => {
    if (force || Date.now() - lastSave > SAVE_EVERY_MS) {
      lastSave = Date.now();
      stored = saveJob(job);
    }
    onProgress(stored);
  };
  const worker = async (): Promise<void> => {
    while (!token.stop) {
      const i = queue.shift();
      if (i === undefined) {
        return;
      }
      job.pending.push(i);
      if (i >= job.next) {
        job.next = i + 1;
      }
      try {
        const r = await scanner.scanSite(sites[i], job.subjects, job.includeLibraries);
        r.rows.forEach((c) => job.rows.push([i, ...c]));
        r.unread.forEach((c) => job.unread.push([i, ...c]));
        // "No access" / "not found" is about that one site; anything else in a row suggests the connection or sign-in.
        const other = r.unread.filter((u) => ['401', '403', '404'].indexOf(u[4]) < 0);
        failsInARow = other.length ? failsInARow + 1 : 0;
        if (failsInARow >= MAX_FAILS_IN_A_ROW) {
          token.stop = true;
          token.reason = `Paused after ${MAX_FAILS_IN_A_ROW} sites in a row could not be read (last: ${other[other.length - 1][5] || 'no details'}). Check your connection or sign-in, then Resume.`;
        }
      } catch (e) {
        // Cancelled, or something unexpected: not a site problem, so it stays pending for the resume.
        token.stop = true;
        if (!(e instanceof CancelledError)) {
          token.reason = `Paused: ${(e as Error).message || String(e)}`;
        }
        persist(true);
        return;
      }
      job.pending = job.pending.filter((x) => x !== i);
      persist(false);
      if (job.speed === 'gentle' && !token.stop && pauseMs > 0) {
        await new Promise((r) => setTimeout(r, pauseMs));
      }
    }
  };
  await Promise.all(job.speed === 'gentle' ? [worker()] : [worker(), worker()]);
  job.finished = !queue.length && !job.pending.length;
  persist(true);
}
