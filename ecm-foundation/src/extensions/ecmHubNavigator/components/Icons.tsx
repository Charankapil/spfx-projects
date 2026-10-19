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
