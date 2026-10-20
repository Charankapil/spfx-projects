import * as React from 'react';

/** Inline stroke icons: no icon font to load on every page the navigator runs on. */
const base = {
  fill: 'none',
  stroke: 'currentColor',
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true
};

export const SearchIcon = (): JSX.Element => (
  <svg width="16" height="16" viewBox="0 0 24 24" strokeWidth={2} {...base}>
    <circle cx="11" cy="11" r="7" />
    <path d="M20 20l-3.5-3.5" />
  </svg>
);

export const ChevronDown = (): JSX.Element => (
  <svg width="12" height="12" viewBox="0 0 24 24" strokeWidth={2.5} {...base}>
    <path d="M6 9l6 6 6-6" />
  </svg>
);

export const ChevronUp = (): JSX.Element => (
  <svg width="12" height="12" viewBox="0 0 24 24" strokeWidth={2.5} {...base}>
    <path d="M18 15l-6-6-6 6" />
  </svg>
);

export const CloseIcon = (): JSX.Element => (
  <svg width="14" height="14" viewBox="0 0 24 24" strokeWidth={2.4} {...base}>
    <path d="M6 6l12 12M18 6L6 18" />
  </svg>
);

export const MenuIcon = (): JSX.Element => (
  <svg width="18" height="18" viewBox="0 0 24 24" strokeWidth={2.2} {...base}>
    <path d="M4 7h16M4 12h16M4 17h16" />
  </svg>
);

export const LockIcon = (): JSX.Element => (
  <svg width="10" height="10" viewBox="0 0 24 24" strokeWidth={2.6} {...base}>
    <rect x="5" y="11" width="14" height="10" rx="2" />
    <path d="M8 11V7a4 4 0 0 1 8 0v4" />
  </svg>
);

export const ShieldIcon = (): JSX.Element => (
  <svg width="14" height="14" viewBox="0 0 24 24" strokeWidth={2} {...base}>
    <path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z" />
  </svg>
);

export const ArrowRight = (): JSX.Element => (
  <svg width="14" height="14" viewBox="0 0 24 24" strokeWidth={2.2} {...base}>
    <path d="M5 12h14M13 6l6 6-6 6" />
  </svg>
);

export const PlusIcon = (): JSX.Element => (
  <svg width="16" height="16" viewBox="0 0 24 24" strokeWidth={2.4} {...base}>
    <path d="M12 5v14M5 12h14" />
  </svg>
);

/** One line icon per functional hub, chosen from the hub title. */
const HUB_ICON_PATHS: { match: RegExp; paths: string[] }[] = [
  { match: /shared|enterprise services/i, paths: ['M4 4h6v6H4z', 'M14 4h6v6h-6z', 'M4 14h6v6H4z', 'M14 14h6v6h-6z'] },
  { match: /corporate|people|hr/i, paths: ['M4 8h16v11H4z', 'M9 8V5h6v3', 'M4 13h16'] },
  { match: /financ|tax|account/i, paths: ['M4 20h16', 'M7 16V10', 'M12 16V6', 'M17 16v-4'] },
  { match: /tech|security|it\b/i, paths: ['M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z', 'M9 12l2 2 4-4'] },
  { match: /operation|network/i, paths: ['M12 10v11', 'M9 21l3-11 3 11', 'M8.5 6.5a5 5 0 0 1 7 0', 'M6 4a8.5 8.5 0 0 1 12 0', 'M12 10a1 1 0 1 0 0-.01'] },
  { match: /real estate|infrastructure|property/i, paths: ['M4 21V9l8-5 8 5v12', 'M9 21v-6h6v6', 'M4 21h16'] },
  { match: /supply|sustain/i, paths: ['M5 19c0-8 6-13 14-14-1 8-6 14-14 14z', 'M5 19l7-7'] }
];

export const HubIcon = (props: { title: string; size?: number }): JSX.Element => {
  const icon = HUB_ICON_PATHS.filter((i) => i.match.test(props.title))[0];
  const paths = icon ? icon.paths : ['M12 3l9 5-9 5-9-5z', 'M3 13l9 5 9-5'];
  const size = props.size || 16;
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" strokeWidth={1.9} {...base}>
      {paths.map((d) => (
        <path key={d} d={d} />
      ))}
    </svg>
  );
};
