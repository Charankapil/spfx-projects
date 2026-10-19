import { NavItemType } from '../../../extensions/ecmHubNavigator/models/INavigation';

export interface ISeedRow {
  title: string;
  type: NavItemType;
  parentHub: string;
  /** Server-relative; made absolute with the tenant's host when seeded. */
  url: string;
  description: string;
  order: number;
  restricted: boolean;
}

/**
 * The hub-and-spoke model from the ECM strategy: six functional hubs, their
 * department sites and hub resources, the shared enterprise services and two
 * header links. Site URLs are placeholders following /sites/<name>; the setup
 * web part's link check shows which ones need correcting in the list.
 */
const ROWS: [string, NavItemType, string, string, string, number, boolean][] = [
  ['Corporate Services', 'Hub', '', '/sites/corporate-services', 'Stricter access reviews for HR and Legal; confidentiality labels on high-value records.', 10, false],
  ['HR', 'Department site', 'Corporate Services', '/sites/hr', 'People policies and employee records', 10, true],
  ['Legal', 'Department site', 'Corporate Services', '/sites/legal', 'Legal matters and privileged content', 20, true],
  ['Public Affairs', 'Department site', 'Corporate Services', '/sites/public-affairs', 'Communications and government relations', 30, false],
  ['Sales & Marketing', 'Department site', 'Corporate Services', '/sites/sales-marketing', 'Campaigns, proposals and brand assets', 40, false],
  ['Product & Strategy', 'Department site', 'Corporate Services', '/sites/product-strategy', 'Product plans and strategy documents', 50, false],
  ['Policies and procedures', 'Hub resource', 'Corporate Services', '/sites/corporate-services/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Corporate Services', '/sites/corporate-services/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Corporate Services', '/sites/corporate-services/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Corporate Services', '/sites/corporate-services/SitePages/News.aspx', '', 40, false],
  ['Finance', 'Hub', '', '/sites/finance-hub', 'Regular access reviews for financial, tax and audit evidence; visitors kept restricted.', 20, false],
  ['Finance', 'Department site', 'Finance', '/sites/finance', 'Planning, budgets and reporting', 10, false],
  ['Accounting', 'Department site', 'Finance', '/sites/accounting', 'Close, ledgers and reconciliations', 20, false],
  ['Tax', 'Department site', 'Finance', '/sites/tax', 'Tax filings and supporting evidence', 30, true],
  ['Global Internal Audit', 'Department site', 'Finance', '/sites/internal-audit', 'Audit plans and evidence', 40, true],
  ['Global Business Services', 'Department site', 'Finance', '/sites/gbs', 'Shared-service processes', 50, false],
  ['Policies and procedures', 'Hub resource', 'Finance', '/sites/finance-hub/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Finance', '/sites/finance-hub/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Finance', '/sites/finance-hub/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Finance', '/sites/finance-hub/SitePages/News.aspx', '', 40, false],
  ['Technology & Security', 'Hub', '', '/sites/technology-security', 'Tightly restricted membership, no broad external sharing, sensitivity labels and quarterly reviews.', 30, false],
  ['IT', 'Department site', 'Technology & Security', '/sites/it', 'Services, architecture and standards', 10, false],
  ['Information Security', 'Department site', 'Technology & Security', '/sites/infosec', 'Cyber policies and incident records', 20, true],
  ['Corporate Security', 'Department site', 'Technology & Security', '/sites/corporate-security', 'Physical security and investigations', 30, true],
  ['Policies and procedures', 'Hub resource', 'Technology & Security', '/sites/technology-security/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Technology & Security', '/sites/technology-security/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Technology & Security', '/sites/technology-security/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Technology & Security', '/sites/technology-security/SitePages/News.aspx', '', 40, false],
  ['Operations', 'Hub', '', '/sites/operations', 'Metadata for location, asset type and criticality, so operationally critical content is prioritized for recovery.', 40, false],
  ['BE', 'Department site', 'Operations', '/sites/be', 'Department documents and processes', 10, false],
  ['Network Operations', 'Department site', 'Operations', '/sites/network-operations', 'Network procedures and operational records', 20, false],
  ['Operational Health & Safety', 'Department site', 'Operations', '/sites/ohs', 'Safety programs, incidents and inspections', 30, false],
  ['Policies and procedures', 'Hub resource', 'Operations', '/sites/operations/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Operations', '/sites/operations/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Operations', '/sites/operations/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Operations', '/sites/operations/SitePages/News.aspx', '', 40, false],
  ['Real Estate & Infrastructure', 'Hub', '', '/sites/real-estate', 'Document types for leases, property and construction records; retention set by legal and operational need.', 50, false],
  ['Leasing', 'Department site', 'Real Estate & Infrastructure', '/sites/leasing', 'Lease agreements, renewals and amendments', 10, false],
  ['Property Management', 'Department site', 'Real Estate & Infrastructure', '/sites/property-management', 'Site and property records', 20, false],
  ['Development & Construction', 'Department site', 'Real Estate & Infrastructure', '/sites/dev-construction', 'Build projects and construction records', 30, false],
  ['Engineering', 'Department site', 'Real Estate & Infrastructure', '/sites/engineering', 'Drawings, specifications and engineering records', 40, false],
  ['Policies and procedures', 'Hub resource', 'Real Estate & Infrastructure', '/sites/real-estate/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Real Estate & Infrastructure', '/sites/real-estate/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Real Estate & Infrastructure', '/sites/real-estate/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Real Estate & Infrastructure', '/sites/real-estate/SitePages/News.aspx', '', 40, false],
  ['Supply Chain & Sustainability', 'Hub', '', '/sites/supply-chain-sustainability', 'Dedicated vendor groups where external sharing is needed; supplier records governed by retention.', 60, false],
  ['Supply Chain', 'Department site', 'Supply Chain & Sustainability', '/sites/supply-chain', 'Supplier lifecycle and procurement', 10, false],
  ['Sustainability', 'Department site', 'Supply Chain & Sustainability', '/sites/sustainability', 'Sustainability programs and reporting', 20, false],
  ['Policies and procedures', 'Hub resource', 'Supply Chain & Sustainability', '/sites/supply-chain-sustainability/SitePages/Policies.aspx', '', 10, false],
  ['Approved templates', 'Hub resource', 'Supply Chain & Sustainability', '/sites/supply-chain-sustainability/Templates', '', 20, false],
  ['Records and retention guidance', 'Hub resource', 'Supply Chain & Sustainability', '/sites/supply-chain-sustainability/SitePages/Records-and-retention.aspx', '', 30, false],
  ['Hub news', 'Hub resource', 'Supply Chain & Sustainability', '/sites/supply-chain-sustainability/SitePages/News.aspx', '', 40, false],
  ['Policy Center', 'Shared service', '', '/sites/policy-center', 'Enterprise policies and standards', 10, false],
  ['Knowledge Center', 'Shared service', '', '/sites/knowledge-center', 'How-to guidance and FAQs', 20, false],
  ['Vendor Portal', 'Shared service', '', '/sites/vendor-portal', 'Supplier collaboration space', 30, true],
  ['Template Center', 'Shared service', '', '/sites/template-center', 'Approved document templates', 40, false],
  ['Records Center', 'Shared service', '', '/sites/records-center', 'File plan and records guidance', 50, false],
  ['Governance Center', 'Shared service', '', '/sites/governance-center', 'ECM roles, rules and scorecards', 60, false],
  ['Request a site', 'Header link', '', '/sites/ecm-hub/SitePages/Request-a-site.aspx', '', 10, false],
  ['Policies and records', 'Header link', '', '/sites/policy-center', '', 20, false]
];

export const NAVIGATION_SEED: ISeedRow[] = ROWS.map(([title, type, parentHub, url, description, order, restricted]) => ({
  title,
  type,
  parentHub,
  url,
  description,
  order,
  restricted
}));
