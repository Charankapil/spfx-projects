/**
 * The ECM palette, shared by the navigator, the web parts and the tenant
 * theme in provisioning/theme/ecm-theme.json. Keep the three in step: the
 * tenant theme colours SharePoint's own chrome, while these values colour the
 * parts SharePoint's theme engine never reaches (the navigator band and the
 * per-hub accents).
 */
export const ECM_PALETTE = {
  navy: '#0E2841',
  deep: '#071B2E',
  blue: '#0067B8',
  sky: '#5CC8FF',
  ink: '#1F2B3A',
  body: '#4A5868',
  muted: '#5B6673',
  line: '#DCE6F0',
  surface: '#FFFFFF',
  canvas: '#F3F5F8',
  amber: '#F2A541'
};

/**
 * Default accent per functional hub, keyed by hub title. A hub row in the
 * navigation list can override its colour; these cover the seven hubs in the
 * strategy so a freshly seeded list already looks right.
 */
export const HUB_COLORS: { [hubTitle: string]: string } = {
  'Corporate Services': '#0067B8',
  Finance: '#8A5A12',
  'Technology & Security': '#00756A',
  Operations: '#5C2D91',
  'Real Estate & Infrastructure': '#B5470E',
  'Supply Chain & Sustainability': '#4E7A16',
  'Shared Enterprise Services': '#2C3E55'
};

const FALLBACK_HUB_COLORS = ['#0067B8', '#8A5A12', '#00756A', '#5C2D91', '#B5470E', '#4E7A16', '#2C3E55'];

export function hubColor(title: string, explicit: string | undefined, index: number): string {
  if (explicit && /^#[0-9a-f]{6}$/i.test(explicit.trim())) {
    return explicit.trim();
  }
  return HUB_COLORS[title] || FALLBACK_HUB_COLORS[index % FALLBACK_HUB_COLORS.length];
}

/**
 * Publishes the palette as CSS custom properties on <html> so any ECM
 * component (or a page's own CSS) can use var(--ecm-navy) and friends
 * without importing this module.
 */
export function applyEcmCssVariables(doc: Document = document): void {
  const root = doc.documentElement;
  Object.keys(ECM_PALETTE).forEach((key) => {
    root.style.setProperty(`--ecm-${key}`, (ECM_PALETTE as { [k: string]: string })[key]);
  });
}
