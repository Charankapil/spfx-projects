/**
 * Chart colours. Series slots 1-3 of the validated reference palette (checked
 * with the dataviz validator against white: CVD ΔE 9.2 all-pairs; slot 3 is
 * under 3:1 contrast, so every chart that uses it carries direct labels and a
 * table view). Status colours are reserved for good / warning / critical and
 * always ship with an icon and a label.
 */
export const SERIES = {
  score: '#2a78d6',
  metadata: '#eb6834',
  labelled: '#1baf7a'
};

export const STATUS_COLOR = {
  good: '#0ca30c',
  warning: '#fab219',
  critical: '#d03b3b'
};

export const INK = {
  primary: '#1E2227',
  secondary: '#4B5260',
  muted: '#6B717C',
  grid: '#E6E8EC',
  axis: '#C8CDD3',
  deEmphasis: '#A3A9B2',
  deltaGood: '#006300',
  deltaBad: '#B42318'
};
