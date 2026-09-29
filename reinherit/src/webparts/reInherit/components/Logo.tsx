import * as React from 'react';

/** The ReInherit mark: a hierarchy inside a "restore" arrow (same as assets/reinherit-logo.svg). */
export const Logo: React.FC<{ size?: number; className?: string }> = ({ size = 40, className }) => (
  <svg viewBox="0 0 96 96" width={size} height={size} className={className} role="img" aria-label="ReInherit">
    <defs>
      <linearGradient id="reinherit-logo-bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stopColor="#0d9488" />
        <stop offset="1" stopColor="#4338ca" />
      </linearGradient>
    </defs>
    <rect width="96" height="96" rx="22" fill="url(#reinherit-logo-bg)" />
    <path d="M25.4 72.6A32 32 0 1 1 71.3 71.9" fill="none" stroke="#ffffff" strokeWidth="6" strokeLinecap="round" />
    <path d="M77.3 77.9L65.3 65.9L63.5 80.1Z" fill="#ffffff" stroke="#ffffff" strokeWidth="2.5" strokeLinejoin="round" />
    <path
      d="M48 42V49M37 54V49H59V54"
      fill="none"
      stroke="#ffffff"
      strokeWidth="4"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
    <circle cx="48" cy="36" r="7" fill="#ffffff" />
    <circle cx="37" cy="60" r="6" fill="#ffffff" />
    <circle cx="59" cy="60" r="6" fill="#a7f3d0" />
  </svg>
);
