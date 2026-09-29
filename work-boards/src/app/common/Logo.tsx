import * as React from 'react';

/** The Work Boards mark: three kanban lanes over a progress bar. Same drawing as docs/brand/work-boards-logo.svg. */
export function Logo(props: { size?: number }): JSX.Element {
  const size = props.size || 24;
  const gradientId = React.useMemo(() => 'wb-logo-' + Math.random().toString(36).slice(2, 8), []);
  return (
    <svg width={size} height={size} viewBox="0 0 96 96" aria-hidden="true" focusable="false" style={{ flex: 'none', display: 'block' }}>
      <defs>
        <linearGradient id={gradientId} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#2f4bd0" />
          <stop offset="1" stopColor="#6a3fd6" />
        </linearGradient>
      </defs>
      <rect width="96" height="96" rx="22" fill={`url(#${gradientId})`} />
      <g fill="#ffffff">
        <rect x="17" y="18" width="18" height="24" rx="5" />
        <rect x="17" y="46" width="18" height="16" rx="5" fillOpacity="0.55" />
        <rect x="39" y="18" width="18" height="38" rx="5" />
        <rect x="61" y="18" width="18" height="14" rx="5" />
        <rect x="61" y="36" width="18" height="22" rx="5" fillOpacity="0.55" />
        <rect x="17" y="68" width="62" height="10" rx="5" fillOpacity="0.25" />
      </g>
      <rect x="17" y="68" width="40" height="10" rx="5" fill="#00c875" />
    </svg>
  );
}
