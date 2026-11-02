import { IDataProvider } from '../providers/IDataProvider';
import { IExportSection } from '../export/tabular';
import { toExportTable } from './components/DataGrid';
import {
  accountColumns, agentColumns, appColumns, connectorColumns, credentialColumns, dlpColumns, envColumns, findingColumns, flattenRoles,
  flowColumns, groupColumns, inactiveUserColumns, issueColumns, makerColumns, messageColumns, mfaColumns, roleColumns, siteColumns,
  skuColumns, teamActivityColumns
} from './columns';
import {
  adoptionSummary, agentSummary, appSummary, connectorUsage, envNameMap, envRows, envSummary, filterInventory, findingSummary, findings,
  flowSummary, identitySummary, inactiveLicensed, IPPData, makerSummary, siteSummary, skuRows, teamsSummary, unusedCost
} from './derive';
import { PageId } from './AppContext';
import { daysSince, fmtNumber } from '../logic/format';
import { buildMakers } from '../providers/LiveDataProvider';

export interface IPageDef {
  id: PageId;
  title: string;
  group: 'Overview' | 'Power Platform' | 'Microsoft 365' | 'Reports';
  icon: string;
  description: string;
  /** Builds the export for this page; undefined for pages without data. */
  build?: (p: IDataProvider, envFilter: string[], prices: Record<string, number>) => Promise<IExportSection>;
}

export async function loadPP(p: IDataProvider, envFilter: string[]): Promise<IPPData> {
  const [inv, makers] = await Promise.all([p.getInventory(), p.getMakers()]);
  const filtered = filterInventory(inv, envFilter);
  // Per-maker counts are recomputed so they reflect only the selected environments.
  return { inv: filtered, makers: envFilter.length ? { ...makers, makers: buildMakers(filtered, makers.users) } : makers };
}

async function optional<T>(work: Promise<T>): Promise<T | undefined> {
  try {
    return await work;
  } catch {
    return undefined;
  }
}

const short = (s: { label: string; value: string }[]): { label: string; value: string }[] => s.map((x) => ({ label: x.label, value: x.value }));

export const PAGES: IPageDef[] = [
  {
    id: 'overview', title: 'Overview', group: 'Overview', icon: 'ViewDashboard',
    description: 'Tenant-wide Power Platform inventory, growth, top risks and service health.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      const dlp = await optional(p.getDlpPolicies());
      const list = findings(pp, dlp, undefined, undefined);
      return {
        title: 'Overview',
        summary: [
          { label: 'Environments', value: fmtNumber(pp.inv.environments.length) },
          { label: 'Apps', value: fmtNumber(pp.inv.apps.length) },
          { label: 'Cloud flows', value: fmtNumber(pp.inv.flows.length) },
          { label: 'Agents', value: fmtNumber(pp.inv.agents.length) },
          { label: 'Makers', value: fmtNumber(pp.makers.makers.length) }
        ].concat(short(findingSummary(list))),
        tables: [toExportTable('Top findings', findingColumns, list.slice(0, 200))]
      };
    }
  },
  {
    id: 'environments', title: 'Environments', group: 'Power Platform', icon: 'Globe',
    description: 'Every environment with its type, region, Dataverse, Managed Environment state, DLP coverage, capacity and resource counts.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      const dlp = await optional(p.getDlpPolicies());
      const rows = envRows(pp.inv, dlp);
      return {
        title: 'Environments',
        summary: short(envSummary(rows, !!dlp)),
        tables: [
          toExportTable('Environments', envColumns, rows),
          {
            name: 'Environment errors',
            columns: [{ header: 'Environment' }, { header: 'Area' }, { header: 'Message', width: 80 }],
            rows: pp.inv.errors.map((e) => [e.environmentName, e.area, e.message])
          }
        ]
      };
    }
  },
  {
    id: 'apps', title: 'Apps', group: 'Power Platform', icon: 'AppIconDefault',
    description: 'Canvas and model-driven apps with owners, connectors, sharing and age. Managed model-driven apps installed by Microsoft or ISV solutions are not counted.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      return { title: 'Apps', summary: short(appSummary(pp.inv.apps, pp.makers)), tables: [toExportTable('Apps', appColumns(pp.makers.users), pp.inv.apps)] };
    }
  },
  {
    id: 'flows', title: 'Cloud flows', group: 'Power Platform', icon: 'Flow',
    description: 'Cloud flows by state, trigger, connectors and owner. Open a flow to see its recent runs.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      return { title: 'Cloud flows', summary: short(flowSummary(pp.inv.flows, pp.makers)), tables: [toExportTable('Cloud flows', flowColumns(pp.makers.users), pp.inv.flows)] };
    }
  },
  {
    id: 'agents', title: 'Copilot agents', group: 'Power Platform', icon: 'Robot',
    description: 'Copilot Studio agents in every Dataverse environment, with publishing and authentication status.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      return { title: 'Copilot agents', summary: short(agentSummary(pp.inv.agents, pp.makers)), tables: [toExportTable('Agents', agentColumns(pp.makers.users), pp.inv.agents)] };
    }
  },
  {
    id: 'makers', title: 'Makers', group: 'Power Platform', icon: 'People',
    description: 'Everyone who owns an app, flow or agent, with department, status and activity.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      return { title: 'Makers', summary: short(makerSummary(pp.makers)), tables: [toExportTable('Makers', makerColumns, pp.makers.makers)] };
    }
  },
  {
    id: 'connectors', title: 'Connectors & DLP', group: 'Power Platform', icon: 'PlugConnected',
    description: 'Which connectors are used where, how DLP policies classify them, and which environments no policy covers.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      const dlp = await optional(p.getDlpPolicies());
      const usage = connectorUsage(pp.inv.apps, pp.inv.flows, dlp);
      const names = envNameMap(pp.inv.environments);
      return {
        title: 'Connectors & DLP',
        summary: [
          { label: 'Connectors in use', value: fmtNumber(usage.length) },
          { label: 'Premium', value: fmtNumber(usage.filter((u) => u.tier === 'Premium').length) },
          { label: 'Custom', value: fmtNumber(usage.filter((u) => u.tier === 'Custom').length) },
          { label: 'DLP policies', value: dlp ? fmtNumber(dlp.length) : '—' }
        ],
        tables: [toExportTable('Connector usage', connectorColumns, usage)].concat(dlp ? [toExportTable('DLP policies', dlpColumns(names), dlp)] : [])
      };
    }
  },
  {
    id: 'capacity', title: 'Licensing & capacity', group: 'Power Platform', icon: 'Database',
    description: 'Power Platform licences, Dataverse capacity by environment and the cost of unassigned licences.',
    build: async (p, env, prices) => {
      const pp = await loadPP(p, env);
      const skus = await optional(p.getSkus());
      const rows = skuRows((skus || []).filter((s) => s.isPowerPlatform), prices);
      const envs = envRows(pp.inv, undefined).filter((e) => e.hasDataverse);
      return {
        title: 'Licensing & capacity',
        summary: [
          { label: 'Power Platform SKUs', value: fmtNumber(rows.length) },
          { label: 'Unassigned licences', value: fmtNumber(rows.reduce((s, r) => s + r.available, 0)) },
          { label: 'Unassigned cost / month', value: unusedCost(rows) ? fmtNumber(unusedCost(rows)) : '—' },
          { label: 'Dataverse environments', value: fmtNumber(envs.length) }
        ],
        tables: [toExportTable('Power Platform licences', skuColumns, rows), toExportTable('Dataverse capacity', envColumns, envs)]
      };
    }
  },
  {
    id: 'governance', title: 'Governance actions', group: 'Power Platform', icon: 'Shield',
    description: 'One worklist of risks found by the governance rules, each with a recommended action and a link to fix it.',
    build: async (p, env) => {
      const pp = await loadPP(p, env);
      const [dlp, teams, credentials] = await Promise.all([optional(p.getDlpPolicies()), optional(p.getTeams()), optional(p.getAppCredentials())]);
      const list = findings(pp, dlp, teams, credentials);
      return { title: 'Governance actions', summary: short(findingSummary(list)), tables: [toExportTable('Findings', findingColumns, list)] };
    }
  },
  {
    id: 'sites', title: 'SharePoint sites', group: 'Microsoft 365', icon: 'SharepointLogo',
    description: 'Site inventory from the SharePoint usage report: storage, activity, sharing and inactive sites.',
    build: async (p) => {
      const data = await p.getSites();
      return {
        title: 'SharePoint sites',
        summary: short(siteSummary(data)).concat(data.sharing ? [{ label: 'Tenant sharing level', value: data.sharing.sharingCapability }] : []),
        tables: [toExportTable('Sites', siteColumns, data.sites)]
      };
    }
  },
  {
    id: 'teams', title: 'Teams & groups', group: 'Microsoft 365', icon: 'TeamsLogo',
    description: 'Microsoft 365 groups and teams with owners, guests and activity.',
    build: async (p) => {
      const data = await p.getTeams();
      return {
        title: 'Teams & groups',
        summary: short(teamsSummary(data)),
        tables: [toExportTable('Groups', groupColumns, data.groups), toExportTable('Team activity', teamActivityColumns, data.activity)]
      };
    }
  },
  {
    id: 'adoption', title: 'Licensing & adoption', group: 'Microsoft 365', icon: 'Chart',
    description: 'All licences, daily active users per workload, Microsoft 365 Copilot usage and licences assigned to inactive users.',
    build: async (p, _env, prices) => {
      const [data, skus] = await Promise.all([p.getAdoption(), optional(p.getSkus())]);
      const rows = skuRows(skus || [], prices);
      return {
        title: 'Licensing & adoption',
        summary: short(adoptionSummary(data, skus)),
        tables: [
          toExportTable('All licences', skuColumns, rows),
          toExportTable('Inactive licensed users', inactiveUserColumns, inactiveLicensed(data)),
          {
            name: 'Daily active users',
            columns: [{ header: 'Date', isDate: true }, { header: 'Microsoft 365' }, { header: 'Exchange' }, { header: 'Teams' }, { header: 'SharePoint' }, { header: 'OneDrive' }],
            rows: data.counts.map((c) => [c.date, c.office365, c.exchange, c.teams, c.sharePoint, c.oneDrive])
          }
        ]
      };
    }
  },
  {
    id: 'identity', title: 'Identity & roles', group: 'Microsoft 365', icon: 'Permissions',
    description: 'Admin role holders, guests, stale accounts, MFA registration and expiring app credentials.',
    build: async (p) => {
      const data = await p.getIdentity();
      const stale = data.users.filter((u) => u.accountEnabled && (daysSince(u.lastSignIn) === undefined || (daysSince(u.lastSignIn) || 0) > 90));
      return {
        title: 'Identity & roles',
        summary: short(identitySummary(data)),
        tables: [
          toExportTable('Admin roles', roleColumns, flattenRoles(data.roles)),
          toExportTable('Guests', accountColumns, data.users.filter((u) => u.userType === 'Guest')),
          ...(data.signInUnavailable ? [] : [toExportTable('Stale accounts', accountColumns, stale)]),
          ...(data.mfaUnavailable ? [] : [toExportTable('MFA registration', mfaColumns, data.mfa)]),
          toExportTable('App credentials', credentialColumns, data.credentials)
        ]
      };
    }
  },
  {
    id: 'export', title: 'Export centre', group: 'Reports', icon: 'Download',
    description: 'Pick pages and a format to download a combined report. Service health and message center are included as their own section.'
  }
];

/** Service health is shown on the overview; exported as its own section. */
export async function buildHealthSection(p: IDataProvider): Promise<IExportSection> {
  const h = await p.getServiceHealth();
  return {
    title: 'Service health',
    summary: [
      { label: 'Services with issues', value: fmtNumber(h.overviews.filter((o) => o.status !== 'serviceOperational').length) },
      { label: 'Open issues', value: fmtNumber(h.issues.length) }
    ],
    tables: [toExportTable('Open issues', issueColumns, h.issues), toExportTable('Message center', messageColumns, h.messages)]
  };
}
