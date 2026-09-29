/** Label and group colours, close to the ones people know from monday.com. */
export const PALETTE: { name: string; value: string }[] = [
  { name: 'Green', value: '#00c875' },
  { name: 'Dark green', value: '#037f4c' },
  { name: 'Orange', value: '#fdab3d' },
  { name: 'Red', value: '#e2445c' },
  { name: 'Dark red', value: '#bb3354' },
  { name: 'Blue', value: '#0086c0' },
  { name: 'Light blue', value: '#579bfc' },
  { name: 'Purple', value: '#a25ddc' },
  { name: 'Dark purple', value: '#784bd1' },
  { name: 'Pink', value: '#ff5ac4' },
  { name: 'Yellow', value: '#ffcb00' },
  { name: 'Brown', value: '#7f5347' },
  { name: 'Grey', value: '#c4c4c4' },
  { name: 'Dark grey', value: '#757575' }
];

export const DEFAULT_LABEL_COLOR = '#c4c4c4';

/** Pick black or white text for a label background. */
export function textOn(background: string): string {
  const hex = background.replace('#', '');
  if (hex.length !== 6) {
    return '#ffffff';
  }
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  // Perceived luminance (ITU-R BT.601).
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return lum > 0.72 ? '#323338' : '#ffffff';
}
