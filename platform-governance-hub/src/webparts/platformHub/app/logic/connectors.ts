import { ConnectorTier, IConnectorRef } from '../models';

/**
 * Connectors that require a premium (per-user / per-app / pay-as-you-go) licence.
 * Used only when the API response does not state the tier itself.
 */
const KNOWN_PREMIUM = [
  'shared_commondataserviceforapps', 'shared_commondataservice', 'shared_sql', 'shared_azureblob',
  'shared_azuretables', 'shared_azurequeues', 'shared_servicebus', 'shared_eventhubs', 'shared_documentdb',
  'shared_azuread', 'shared_azureautomation', 'shared_azuredatafactory', 'shared_keyvault', 'shared_salesforce',
  'shared_sapodata', 'shared_saperp', 'shared_oracle', 'shared_db2', 'shared_mysql', 'shared_postgresql',
  'shared_servicenow', 'shared_jira', 'shared_adobesign', 'shared_docusign', 'shared_webcontents',
  'shared_uiflow', 'shared_flowrpa', 'shared_dynamicscrmonline', 'shared_dynamicsax', 'shared_azureopenai',
  'shared_cognitiveservicestextanalytics', 'shared_filesystem', 'shared_ftp', 'shared_sftp', 'shared_sftpwithssh',
  'shared_http', 'http', 'shared_httpwithazuread', 'shared_webcontentsv2', 'shared_aibuilder'
];

const FRIENDLY: Record<string, string> = {
  shared_sharepointonline: 'SharePoint',
  shared_office365: 'Office 365 Outlook',
  shared_office365users: 'Office 365 Users',
  shared_office365groups: 'Office 365 Groups',
  shared_teams: 'Microsoft Teams',
  shared_onedriveforbusiness: 'OneDrive for Business',
  shared_excelonlinebusiness: 'Excel Online (Business)',
  shared_approvals: 'Approvals',
  shared_planner: 'Planner',
  shared_commondataserviceforapps: 'Microsoft Dataverse',
  shared_commondataservice: 'Dataverse (legacy)',
  shared_sql: 'SQL Server',
  shared_azureblob: 'Azure Blob Storage',
  shared_azuread: 'Azure AD',
  shared_salesforce: 'Salesforce',
  shared_servicenow: 'ServiceNow',
  shared_webcontents: 'HTTP with Microsoft Entra ID',
  shared_http: 'HTTP',
  http: 'HTTP',
  shared_flowpush: 'Notifications',
  shared_powerbi: 'Power BI',
  shared_forms: 'Microsoft Forms',
  shared_outlook: 'Outlook.com',
  shared_gmail: 'Gmail',
  shared_twitter: 'X (Twitter)',
  shared_dropbox: 'Dropbox',
  shared_googledrive: 'Google Drive',
  shared_logicflows: 'Logic flows',
  shared_uiflow: 'Desktop flows',
  shared_azureopenai: 'Azure OpenAI',
  shared_keyvault: 'Azure Key Vault',
  shared_onenote: 'OneNote',
  shared_todo: 'Microsoft To Do',
  shared_wordonlinebusiness: 'Word Online (Business)',
  shared_kaizala: 'Kaizala',
  shared_microsoftforms: 'Microsoft Forms'
};

/** "/providers/Microsoft.PowerApps/apis/shared_sharepointonline" -> "shared_sharepointonline" */
export function connectorNameFromId(id: string | undefined): string {
  if (!id) return '';
  const parts = String(id).split('/');
  return parts[parts.length - 1].toLowerCase();
}

/** Custom connectors carry a generated hash suffix, e.g. shared_contoso-20api-5f1e3c4b9a8d7e6f12. */
const CUSTOM_SUFFIX = /-[0-9a-f]{16,}$/;

export function isLikelyCustom(name: string): boolean {
  return CUSTOM_SUFFIX.test(name);
}

export function friendlyConnectorName(name: string, fallback?: string): string {
  if (fallback && fallback.trim()) return fallback;
  if (FRIENDLY[name]) return FRIENDLY[name];
  const stripped = name.replace(/^shared_/, '').replace(CUSTOM_SUFFIX, '').replace(/-20/g, ' ').replace(/[-_]/g, ' ');
  return stripped.charAt(0).toUpperCase() + stripped.slice(1);
}

export function resolveTier(name: string, declared?: string, isCustom?: boolean): ConnectorTier {
  if (isCustom || isLikelyCustom(name)) return 'Custom';
  if (declared && /premium/i.test(declared)) return 'Premium';
  if (declared && /standard/i.test(declared)) return 'Standard';
  return KNOWN_PREMIUM.indexOf(name) >= 0 ? 'Premium' : 'Standard';
}

export function makeConnector(id: string | undefined, displayName?: string, declaredTier?: string, isCustom?: boolean): IConnectorRef | undefined {
  const name = connectorNameFromId(id);
  if (!name) return undefined;
  return { name, displayName: friendlyConnectorName(name, displayName), tier: resolveTier(name, declaredTier, isCustom) };
}

/** De-duplicates by connector name, keeping the first (most informative) entry. */
export function uniqueConnectors(list: (IConnectorRef | undefined)[]): IConnectorRef[] {
  const seen: Record<string, boolean> = {};
  const out: IConnectorRef[] = [];
  list.forEach((c) => {
    if (c && !seen[c.name]) {
      seen[c.name] = true;
      out.push(c);
    }
  });
  return out;
}
