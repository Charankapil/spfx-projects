import { CertificateCredential, tokenRoles } from './auth.js';
import { Config } from './config.js';
import { Logger } from './log.js';
import { Clients } from './runner.js';
import { SiteScanner } from './scanner.js';
import { originOf } from './sharepoint.js';
import { SharePointStore } from './store.js';
import { parsePrincipalsCsv } from './config.js';

/**
 * "Will the scheduled run work?" - each requirement tested on its own, with the fix for the
 * ones that fail. Run it once after setting up the app registration, and again whenever a
 * scheduled run fails.
 */
export async function check(cfg: Config, clients: Clients, log: Logger): Promise<boolean> {
  let ok = true;
  const pass = (m: string): void => log.info(`[ ok ] ${m}`);
  const fail = (m: string): void => {
    ok = false;
    log.error(`[FAIL] ${m}`);
  };
  const note = (m: string): void => log.info(`[note] ${m}`);
  const has = (roles: string[], ...any: string[]): boolean => any.some((a) => roles.indexOf(a) >= 0);

  // 1. Credential
  if (clients.credential instanceof CertificateCredential) {
    const c = clients.credential;
    const days = c.daysLeft();
    if (days < 0) {
      fail(`The certificate expired on ${c.expires.toISOString().substring(0, 10)}. Create a new one and upload it to the app registration.`);
    } else {
      pass(`Certificate read (thumbprint ${c.thumbprint}, expires ${c.expires.toISOString().substring(0, 10)}, ${days} days left).`);
      if (days < 45) {
        note('The certificate expires soon. Replace it before then.');
      }
    }
  } else {
    pass('Federated sign-in (no stored secret).');
  }

  // 2. Microsoft Graph
  let graphOk = false;
  try {
    const roles = tokenRoles(await clients.tokens.getToken(cfg.endpoints.graph));
    graphOk = true;
    pass(`Signed in to Microsoft Graph. Application permissions: ${roles.length ? roles.join(', ') : '(none)'}.`);
    if (!cfg.sites.only.length && !has(roles, 'Sites.Read.All', 'Sites.ReadWrite.All', 'Sites.FullControl.All', 'Sites.Manage.All')) {
      fail('Microsoft Graph permission Sites.Read.All (application) is missing: it is needed to list every site. Add it and grant admin consent, or list the sites in "sites.only".');
    }
    if (!has(roles, 'User.Read.All', 'User.ReadBasic.All', 'Directory.Read.All')) {
      fail('Microsoft Graph permission User.Read.All (application) is missing: it is needed to look up the people in the report.');
    }
    if (cfg.scan.expandGroups && !has(roles, 'GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All')) {
      fail('Microsoft Graph permission GroupMember.Read.All (application) is missing: it is needed to look up groups and memberships. Add it, or set "scan.expandGroups" to false.');
    }
  } catch (e) {
    fail((e as Error).message);
  }

  // 3. SharePoint
  let spOk = false;
  try {
    const roles = tokenRoles(await clients.tokens.getToken(originOf(cfg.report.siteUrl)));
    spOk = true;
    pass(`Signed in to SharePoint. Application permissions: ${roles.length ? roles.join(', ') : '(none)'}.`);
    if (!has(roles, 'Sites.FullControl.All')) {
      if (has(roles, 'Sites.Selected')) {
        note('SharePoint permission is Sites.Selected: only sites the app was granted Full Control on can be read. List them in "sites.only".');
      } else {
        fail('SharePoint permission Sites.FullControl.All (application) is missing. Reading who has access to a site requires it (lower permissions cannot read permissions). Add it under API permissions > SharePoint and grant admin consent.');
      }
    }
  } catch (e) {
    fail((e as Error).message);
  }
  if (!spOk) {
    return false;
  }

  // 4. Report folder
  const store = new SharePointStore(clients.sp, cfg.report.siteUrl, cfg.report.libraryPath, cfg.report.folder);
  let folderOk = false;
  try {
    await store.prepare();
    await store.write('_state/check.txt', Buffer.from(`Written by the access report check on ${new Date().toISOString()}\r\n`, 'utf8'));
    await store.remove('_state/check.txt');
    folderOk = true;
    pass(`Report folder is writable: ${store.description}`);
  } catch (e) {
    fail(`Report folder: ${(e as Error).message}`);
  }

  // 5. People and groups
  let principals: Awaited<ReturnType<typeof clients.graph.resolvePrincipals>> | undefined;
  if (graphOk) {
    try {
      const requests = cfg.principals.slice();
      if (cfg.principalsFile && folderOk) {
        const data = await store.read(cfg.principalsFile);
        if (data) {
          const parsed = parsePrincipalsCsv(Buffer.from(data).toString('utf8'));
          parsed.problems.forEach((p) => fail(p));
          parsed.principals.forEach((p) => requests.push(p));
        } else if (!requests.length) {
          note(`No ${cfg.principalsFile} in the report folder yet. The first run creates an empty one for you to fill in.`);
        }
      }
      if (requests.length) {
        principals = await clients.graph.resolvePrincipals(requests, cfg.scan.expandGroups);
        principals.unresolved.forEach((u) => fail(`${u.type} "${u.input}": ${u.reason}`));
        principals.warnings.forEach((w) => fail(w));
        principals.principals.forEach((p) => log.debug(`${p.label}: ${p.identities.length} claims`));
        if (principals.principals.length) {
          pass(`${principals.principals.length} of ${requests.length} people and groups found in the directory.`);
        }
      }
    } catch (e) {
      fail(`People and groups: ${(e as Error).message}`);
    }
  }

  // 6. Sites
  if (graphOk && !cfg.sites.only.length) {
    try {
      const sites = await clients.graph.listSites(cfg.sites.includeOneDrive);
      if (sites.length) {
        pass(`${sites.length} site collections listed.`);
      } else {
        fail('Microsoft Graph listed no sites.');
      }
    } catch (e) {
      fail(`Listing sites: ${(e as Error).message}`);
    }
  } else if (cfg.sites.only.length) {
    pass(`${cfg.sites.only.length} sites listed in "sites.only".`);
  }

  // 7. A real permission read, on the report site itself.
  if (principals && principals.principals.length) {
    try {
      const scanner = new SiteScanner(clients.sp, principals.principals, {
        includeLibraries: cfg.scan.includeLibraries,
        includeLists: cfg.scan.includeLists,
        includeSubsites: false,
        thorough: cfg.scan.thorough,
        groupsExpanded: cfg.scan.expandGroups && !principals.warnings.length,
        maxSubsites: cfg.scan.maxSubsites
      });
      const rows = await scanner.scanSite({ url: cfg.report.siteUrl, title: 'Report site' });
      pass(`Permissions of the report site were read (${rows.length} places where the listed people and groups have access).`);
    } catch (e) {
      fail(`Reading permissions of the report site failed: ${(e as Error).message}. This usually means Sites.FullControl.All is missing or not consented.`);
    }
  }
  log.info(ok ? 'All checks passed. The scheduled run should work.' : 'Some checks failed. Fix them and run the check again.');
  return ok;
}
