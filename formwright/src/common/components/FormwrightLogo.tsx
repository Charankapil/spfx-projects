import * as React from 'react';

let instance = 0;

/**
 * The Formwright logo (docs/images/formwright-logo.svg), inline so it needs no
 * image file and stays crisp at any size.
 */
export const FormwrightLogo: React.FC<{ size: number; className?: string }> = (props) => {
  // Gradient ids must be unique on the page.
  const id = React.useMemo(() => `fw-logo-${instance++}`, []);
  return (
    <svg viewBox="0 0 96 96" width={props.size} height={props.size} className={props.className} aria-hidden="true" focusable="false">
      <defs>
        <linearGradient id={`${id}-bg`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#4F46E5" />
          <stop offset="1" stopColor="#0D9488" />
        </linearGradient>
        <linearGradient id={`${id}-pen`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor="#FBBF24" />
          <stop offset="1" stopColor="#F59E0B" />
        </linearGradient>
      </defs>
      <rect width="96" height="96" rx="22" fill={`url(#${id}-bg)`} />
      <rect x="18" y="16" width="46" height="62" rx="7" fill="#FFFFFF" />
      <rect x="26" y="26" width="20" height="5" rx="2.5" fill="#4F46E5" />
      <rect x="26" y="35" width="30" height="8" rx="2.5" fill="#E0E7FF" />
      <rect x="26" y="49" width="14" height="5" rx="2.5" fill="#0D9488" />
      <rect x="26" y="58" width="30" height="8" rx="2.5" fill="#CCFBF1" />
      <g transform="rotate(38 70 58)">
        <rect x="64" y="30" width="12" height="34" rx="3" fill={`url(#${id}-pen)`} stroke="#FFFFFF" strokeWidth="2.5" />
        <rect x="64" y="30" width="12" height="7" rx="3" fill="#FFFFFF" />
        <path d="M64 64 h12 l-6 12 z" fill="#FFFFFF" />
        <path d="M67.5 66 h5 l-2.5 5 z" fill="#1E1B4B" />
      </g>
    </svg>
  );
};
