import * as React from 'react';

/** Platform Lens mark: a lens over rising bars. Mirrors assets/logo.svg. */
export const LensLogo: React.FC<{ size?: number }> = ({ size = 32 }) => {
  const id = React.useMemo(() => `pl-bg-${Math.random().toString(36).slice(2, 8)}`, []);
  return (
    <svg width={size} height={size} viewBox="0 0 96 96" aria-hidden="true" focusable="false" style={{ flex: 'none' }}>
      <defs>
        <linearGradient id={id} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#14857f" />
          <stop offset="1" stopColor="#1b4fa8" />
        </linearGradient>
      </defs>
      <rect width="96" height="96" rx="22" fill={`url(#${id})`} />
      <circle cx="42" cy="42" r="22" fill="#ffffff" fillOpacity={0.14} stroke="#ffffff" strokeWidth={7} />
      <rect x="30.5" y="43" width="6" height="10" rx="1.5" fill="#ffffff" />
      <rect x="39" y="37" width="6" height="16" rx="1.5" fill="#ffffff" />
      <rect x="47.5" y="30" width="6" height="23" rx="1.5" fill="#7ff0d6" />
      <line x1="59" y1="59" x2="75" y2="75" stroke="#ffffff" strokeWidth={9} strokeLinecap="round" />
    </svg>
  );
};
