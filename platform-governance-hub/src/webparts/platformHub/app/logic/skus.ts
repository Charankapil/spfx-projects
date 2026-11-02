/** Friendly names for common SKU part numbers (Graph returns only the part number). */
const NAMES: Record<string, string> = {
  POWERAPPS_PER_USER: 'Power Apps Premium',
  POWERAPPS_PER_APP: 'Power Apps per app',
  POWERAPPS_PER_APP_NEW: 'Power Apps per app',
  POWERAPPS_VIRAL: 'Power Apps (trial / viral)',
  POWERAPPS_DEV: 'Power Apps Developer Plan',
  POWERAPPS_PORTALS_LOGIN_T2: 'Power Pages authenticated users',
  POWERAPPS_PORTALS_PAGEVIEW: 'Power Pages anonymous users',
  FLOW_PER_USER: 'Power Automate per user',
  FLOW_PER_PROCESS: 'Power Automate Process',
  FLOW_FREE: 'Power Automate Free',
  FLOW_P2_VIRAL: 'Power Automate (trial)',
  POWERAUTOMATE_ATTENDED_RPA: 'Power Automate Premium',
  POWERAUTOMATE_UNATTENDED_RPA: 'Power Automate unattended RPA',
  POWER_AUTOMATE_PREMIUM: 'Power Automate Premium',
  VIRTUAL_AGENT_BASE: 'Copilot Studio (base)',
  VIRTUAL_AGENT_USL: 'Copilot Studio user licence',
  POWER_VIRTUAL_AGENTS_BASE: 'Copilot Studio (base)',
  CCIBOTS_PRIVPREV_VIRAL: 'Copilot Studio (trial)',
  Microsoft_Copilot_Studio: 'Copilot Studio',
  Microsoft_365_Copilot: 'Microsoft 365 Copilot',
  AI_BUILDER_CAPACITY_ADDON: 'AI Builder capacity',
  CDS_DB_CAPACITY: 'Dataverse database capacity',
  CDS_FILE_CAPACITY: 'Dataverse file capacity',
  CDS_LOG_CAPACITY: 'Dataverse log capacity',
  POWER_BI_PRO: 'Power BI Pro',
  POWER_BI_STANDARD: 'Power BI (free)',
  PBI_PREMIUM_PER_USER: 'Power BI Premium per user',
  SPE_E3: 'Microsoft 365 E3',
  SPE_E5: 'Microsoft 365 E5',
  SPE_F1: 'Microsoft 365 F3',
  ENTERPRISEPACK: 'Office 365 E3',
  ENTERPRISEPREMIUM: 'Office 365 E5',
  STANDARDPACK: 'Office 365 E1',
  O365_BUSINESS_PREMIUM: 'Microsoft 365 Business Standard',
  SPB: 'Microsoft 365 Business Premium',
  O365_BUSINESS_ESSENTIALS: 'Microsoft 365 Business Basic',
  EMS: 'Enterprise Mobility + Security E3',
  EMSPREMIUM: 'Enterprise Mobility + Security E5',
  AAD_PREMIUM: 'Microsoft Entra ID P1',
  AAD_PREMIUM_P2: 'Microsoft Entra ID P2',
  TEAMS_EXPLORATORY: 'Teams Exploratory',
  Microsoft_Teams_Premium: 'Teams Premium',
  DYN365_ENTERPRISE_PLAN1: 'Dynamics 365 Plan',
  DYN365_ENTERPRISE_SALES: 'Dynamics 365 Sales Enterprise',
  DYN365_ENTERPRISE_CUSTOMER_SERVICE: 'Dynamics 365 Customer Service Enterprise',
  DYN365_TEAM_MEMBERS: 'Dynamics 365 Team Members',
  EXCHANGESTANDARD: 'Exchange Online (Plan 1)',
  WINDOWS_STORE: 'Windows Store for Business',
  PROJECT_P1: 'Project Plan 1',
  PROJECTPROFESSIONAL: 'Project Plan 3',
  VISIOCLIENT: 'Visio Plan 2'
};

export function skuFriendlyName(part: string): string {
  if (NAMES[part]) return NAMES[part];
  return part.replace(/_/g, ' ');
}

const PP_PART = /POWERAPPS|POWER_APPS|FLOW|POWERAUTOMATE|POWER_AUTOMATE|VIRTUAL_AGENT|COPILOT_STUDIO|CCIBOTS|CDS_|AI_BUILDER|PORTALS|POWER_PAGES|DYN365|DYNAMICS/i;

/** Standalone Power Platform / Dynamics SKUs (not M365 suites that merely include seeded rights). */
export function isPowerPlatformSku(part: string, servicePlans: string[] = []): boolean {
  if (PP_PART.test(part)) return true;
  // A SKU whose every plan is Power Platform (e.g. renamed SKUs) also counts.
  return servicePlans.length > 0 && servicePlans.every((p) => PP_PART.test(p) || /EXCHANGE_S_FOUNDATION/i.test(p));
}
