export interface IFileCategory {
  /** Key into the CategoryLabels localized string. */
  key: string;
  /** Fluent UI icon name. */
  icon: string;
}

/*
 * Categories are told apart by label and icon only. The bars keep the
 * dashboard's two colours (active / inactive), so a category never takes
 * a colour that could be read as "inactive".
 */
export const CATEGORIES: IFileCategory[] = [
  { key: 'media', icon: 'Video' },
  { key: 'image', icon: 'Photo2' },
  { key: 'archive', icon: 'ZipFolder' },
  { key: 'pdf', icon: 'PDF' },
  { key: 'word', icon: 'WordDocument' },
  { key: 'excel', icon: 'ExcelDocument' },
  { key: 'powerpoint', icon: 'PowerPointDocument' },
  { key: 'design', icon: 'Design' },
  { key: 'mail', icon: 'Mail' },
  { key: 'database', icon: 'Database' },
  { key: 'code', icon: 'Code' }
];

export const OTHER_CATEGORY: IFileCategory = { key: 'other', icon: 'Page' };

const EXTENSIONS: { [key: string]: string[] } = {
  media: [
    'mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm', 'm4v', 'mpg', 'mpeg', 'mts', 'm2ts', '3gp', 'flv', 'vob',
    'mp3', 'wav', 'm4a', 'wma', 'aac', 'ogg', 'flac', 'aiff'
  ],
  image: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'tif', 'tiff', 'webp', 'heic', 'ico', 'emf', 'wmf', 'raw', 'cr2', 'nef', 'dng'],
  archive: ['zip', 'rar', '7z', 'tar', 'gz', 'tgz', 'cab', 'iso', 'img', 'vhd', 'vhdx', 'vmdk', 'sppkg', 'wsp'],
  pdf: ['pdf'],
  word: ['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'odt', 'rtf', 'txt', 'md', 'one', 'onepkg', 'log'],
  excel: ['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'ods', 'tsv'],
  powerpoint: ['ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'pot', 'potx', 'odp'],
  design: ['psd', 'ai', 'indd', 'eps', 'dwg', 'dxf', 'dwf', 'skp', 'vsd', 'vsdx', 'rvt', 'stp', 'step', 'sldprt', 'fig', 'xd'],
  mail: ['msg', 'eml', 'pst', 'ost', 'mbox'],
  database: ['bak', 'mdb', 'accdb', 'db', 'sqlite', 'mdf', 'ldf', 'dmp', 'sql'],
  code: [
    'aspx', 'html', 'htm', 'css', 'scss', 'js', 'ts', 'tsx', 'jsx', 'json', 'xml', 'xsl', 'xslt',
    'ps1', 'psm1', 'cs', 'py', 'java', 'sh', 'bat', 'yml', 'yaml', 'config', 'master', 'webpart', 'dwp', 'exe', 'msi', 'dll'
  ]
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
