import { IFileTypeStat } from '../../models/IFileTypeStat';

export interface IFileCategory {
  key: string;
  label: string;
  color: string;
  /** Text colour for labels drawn on top of `color`. */
  ink: string;
}

const INK_DARK = '#0b0b0b';
const INK_LIGHT = '#ffffff';

function luminance(hex: string): number {
  const channel = (i: number): number => {
    const c = parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5);
}

/** White or near-black, whichever has the higher contrast against the fill. */
function inkFor(fill: string): string {
  const l = luminance(fill);
  const onDark = (1.05) / (l + 0.05);
  const onLight = (l + 0.05) / (luminance(INK_DARK) + 0.05);
  return onDark >= onLight ? INK_LIGHT : INK_DARK;
}

function category(key: string, label: string, color: string): IFileCategory {
  return { key, label, color, ink: inkFor(color) };
}

/*
 * The eight hues are the validated reference categorical palette (every hard
 * check passes against a white page). Each category keeps its hue in every
 * chart so colour follows the file type, never its rank. More kinds of file
 * than eight fold into a grey "Other" rather than getting generated hues.
 */
export const CATEGORIES: IFileCategory[] = [
  category('word', 'Word & text', '#2a78d6'),
  category('powerpoint', 'PowerPoint', '#eb6834'),
  category('code', 'Web & code', '#1baf7a'),
  category('archive', 'Archives', '#eda100'),
  category('image', 'Images', '#e87ba4'),
  category('excel', 'Excel & data', '#008300'),
  category('media', 'Video & audio', '#4a3aa7'),
  category('pdf', 'PDF', '#e34948')
];

export const OTHER_CATEGORY: IFileCategory = category('other', 'Other', '#c3c2b7');

const EXTENSIONS: { [key: string]: string[] } = {
  word: ['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'odt', 'rtf', 'txt', 'md', 'one'],
  powerpoint: ['ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'pot', 'potx', 'odp'],
  code: [
    'aspx', 'html', 'htm', 'css', 'scss', 'js', 'ts', 'tsx', 'jsx', 'json', 'xml', 'xsl', 'xslt',
    'ps1', 'psm1', 'cs', 'py', 'java', 'sql', 'sh', 'bat', 'yml', 'yaml', 'config', 'master', 'webpart', 'dwp'
  ],
  archive: ['zip', 'rar', '7z', 'tar', 'gz', 'tgz', 'cab', 'iso', 'sppkg', 'wsp'],
  image: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'tif', 'tiff', 'webp', 'heic', 'ico', 'emf', 'wmf'],
  excel: ['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'ods', 'tsv'],
  media: ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm', 'm4v', 'mpg', 'mpeg', 'mp3', 'wav', 'm4a', 'wma', 'aac', 'ogg', 'flac'],
  pdf: ['pdf']
};

const BY_EXTENSION: { [ext: string]: IFileCategory } = {};
CATEGORIES.forEach((c) => {
  (EXTENSIONS[c.key] || []).forEach((ext) => {
    BY_EXTENSION[ext] = c;
  });
});

export function categoryOf(extension: string): IFileCategory {
  return BY_EXTENSION[extension.toLowerCase()] || OTHER_CATEGORY;
}

export interface ICategoryTotal {
  category: IFileCategory;
  count: number;
  extensions: IFileTypeStat[];
}

/** Rolls extension counts up into categories, in the fixed category order (Other last). */
export function totalsByCategory(stats: IFileTypeStat[]): ICategoryTotal[] {
  const totals: { [key: string]: ICategoryTotal } = {};
  stats.forEach((s) => {
    const cat = categoryOf(s.extension);
    if (!totals[cat.key]) {
      totals[cat.key] = { category: cat, count: 0, extensions: [] };
    }
    totals[cat.key].count += s.count;
    totals[cat.key].extensions.push(s);
  });
  return CATEGORIES.concat([OTHER_CATEGORY])
    .map((c) => totals[c.key])
    .filter((t): t is ICategoryTotal => !!t && t.count > 0);
}
