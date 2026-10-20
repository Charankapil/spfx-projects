/**
 * Brand and theme for every ECM component.
 *
 * A brand is two colours and a logo: `primary` (buttons, links, the accent
 * line) and `dark` (the hub band, headings). Everything else is derived, so
 * the SharePoint theme ECM Setup applies and the navigator always agree.
 *
 * The defaults follow American Tower's public identity: a cardinal-red mark
 * with charcoal type on white. Confirm the exact values against the official
 * brand guide and enter them in ECM Setup; nothing else needs to change.
 */
export interface IEcmBrand {
  primary: string;
  dark: string;
  /** Absolute or server-relative URL of the logo image; empty shows the text badge. */
  logoUrl: string;
}

export const DEFAULT_BRAND: IEcmBrand = {
  primary: '#C91F3C',
  dark: '#1E2227',
  logoUrl: ''
};

export const ECM_THEME_NAME = 'American Tower ECM';

/** Neutral greys shared by every component (cool, to sit with the charcoal). */
export const NEUTRALS = {
  ink: '#1E2227',
  body: '#4B5260',
  muted: '#5F6673',
  line: '#E2E5EA',
  soft: '#F4F5F7',
  canvas: '#F7F8FA',
  white: '#FFFFFF'
};

/**
 * Default accent per functional hub, keyed by hub title. Red is kept for the
 * brand itself, so no hub uses it. A hub row in the navigation list can
 * override its colour.
 */
export const HUB_COLORS: { [hubTitle: string]: string } = {
  'Corporate Services': '#1F5FA8',
  Finance: '#8C6210',
  'Technology & Security': '#0B7468',
  Operations: '#5E3FA3',
  'Real Estate & Infrastructure': '#B4510F',
  'Supply Chain & Sustainability': '#4A7A1C',
  'Shared Services': '#4A5568',
  'Shared Enterprise Services': '#4A5568'
};

const FALLBACK_HUB_COLORS = ['#1F5FA8', '#8C6210', '#0B7468', '#5E3FA3', '#B4510F', '#4A7A1C', '#4A5568'];

export function hubColor(title: string, explicit: string | undefined, index: number): string {
  if (explicit && isHexColor(explicit)) {
    return explicit.trim();
  }
  return HUB_COLORS[title] || FALLBACK_HUB_COLORS[index % FALLBACK_HUB_COLORS.length];
}

export function isHexColor(value: string | undefined): boolean {
  return !!value && /^#[0-9a-f]{6}$/i.test(value.trim());
}

/** A brand with any missing or malformed value replaced by the default. */
export function resolveBrand(partial: Partial<IEcmBrand> | undefined): IEcmBrand {
  const primary = (partial?.primary || '').trim();
  const dark = (partial?.dark || '').trim();
  return {
    primary: isHexColor(primary) ? primary : DEFAULT_BRAND.primary,
    dark: isHexColor(dark) ? dark : DEFAULT_BRAND.dark,
    logoUrl: (partial?.logoUrl || '').trim()
  };
}

function toRgb(hex: string): [number, number, number] {
  const n = parseInt(hex.replace('#', ''), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function toHex(rgb: number[]): string {
  return '#' + rgb.map((c) => Math.max(0, Math.min(255, Math.round(c))).toString(16)).map((h) => (h.length < 2 ? '0' + h : h)).join('').toUpperCase();
}

/** Mixes `hex` towards `target` by `amount` (0 = unchanged, 1 = target). */
export function mix(hex: string, target: string, amount: number): string {
  const a = toRgb(hex);
  const b = toRgb(target);
  return toHex(a.map((c, i) => c + (b[i] - c) * amount));
}

/**
 * The SharePoint theme for a brand, in the 27-slot format "Change the look"
 * uses. Theme shades follow the ratios of Microsoft's theme generator closely
 * enough that hover and pressed states stay legible.
 */
export function buildSharePointTheme(brand: IEcmBrand): { name: string; isInverted: boolean; palette: { [slot: string]: string } } {
  const p = brand.primary;
  return {
    name: ECM_THEME_NAME,
    isInverted: false,
    palette: {
      themePrimary: p,
      themeLighterAlt: mix(p, '#FFFFFF', 0.95),
      themeLighter: mix(p, '#FFFFFF', 0.82),
      themeLight: mix(p, '#FFFFFF', 0.68),
      themeTertiary: mix(p, '#FFFFFF', 0.4),
      themeSecondary: mix(p, '#FFFFFF', 0.12),
      themeDarkAlt: mix(p, '#000000', 0.1),
      themeDark: mix(p, '#000000', 0.24),
      themeDarker: mix(p, '#000000', 0.44),
      neutralLighterAlt: '#FAFAFB',
      neutralLighter: NEUTRALS.soft,
      neutralLight: '#EAECEF',
      neutralQuaternaryAlt: NEUTRALS.line,
      neutralQuaternary: '#D7DBE0',
      neutralTertiaryAlt: '#C8CDD3',
      neutralTertiary: '#A3A9B2',
      neutralSecondary: NEUTRALS.muted,
      neutralSecondaryAlt: NEUTRALS.muted,
      neutralPrimaryAlt: '#353A42',
      neutralPrimary: NEUTRALS.ink,
      neutralDark: mix(brand.dark, '#000000', 0.2),
      black: mix(brand.dark, '#000000', 0.45),
      white: NEUTRALS.white,
      primaryBackground: NEUTRALS.white,
      primaryText: NEUTRALS.ink,
      accent: p
    }
  };
}

/**
 * Publishes the brand as CSS custom properties on <html>, so the web parts
 * and any page CSS can use var(--ecm-primary) and friends.
 */
export function applyEcmCssVariables(brand: IEcmBrand, doc: Document = document): void {
  const root = doc.documentElement;
  const vars: { [name: string]: string } = {
    primary: brand.primary,
    'primary-hover': mix(brand.primary, '#000000', 0.12),
    'primary-soft': mix(brand.primary, '#FFFFFF', 0.92),
    dark: brand.dark,
    ...NEUTRALS
  };
  Object.keys(vars).forEach((key) => root.style.setProperty(`--ecm-${key}`, vars[key]));
}
