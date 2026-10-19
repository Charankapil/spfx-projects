/**
 * Metadata arrives in a handful of shapes depending on the column type and the
 * API: plain strings, "1;#Value" lookups, taxonomy objects or "L0|#0guid|Label"
 * search strings, user arrays. Everything is flattened to display labels here
 * so grouping never has to care where a value came from.
 */

interface ILabelled {
  Label?: string;
  lookupValue?: string;
  title?: string;
  Title?: string;
  Description?: string;
  Url?: string;
  value?: string;
}

function fromObject(value: ILabelled): string {
  return value.Label || value.lookupValue || value.title || value.Title || value.Description || value.Url || value.value || '';
}

const TAXONOMY_LABEL = /L0\|#0?[0-9a-f-]+\|([^;]+)/gi;
const TAXONOMY_GP = /GP0\|#/i;

function fromString(value: string, splitMulti: boolean): string[] {
  const text = value.trim();
  if (!text) {
    return [];
  }
  // Managed metadata mapped through ows_taxId_* crawled properties.
  if (TAXONOMY_GP.test(text) || /L0\|#/i.test(text)) {
    const labels: string[] = [];
    let match: RegExpExecArray | null;
    TAXONOMY_LABEL.lastIndex = 0;
    while ((match = TAXONOMY_LABEL.exec(text)) !== null) {
      labels.push(match[1]);
    }
    if (labels.length) {
      return labels;
    }
  }
  // Classic lookup / multi-choice encoding: ";#a;#b;#" or "1;#Label;#2;#Other".
  if (text.indexOf(';#') >= 0) {
    return text
      .split(';#')
      .map((part) => part.trim())
      .filter((part) => part && !/^\d+$/.test(part));
  }
  // Search returns multi-valued managed properties joined with ";".
  if (splitMulti && text.indexOf(';') >= 0) {
    return text
      .split(';')
      .map((part) => part.trim())
      .filter(Boolean);
  }
  return [text];
}

export function toValues(raw: unknown, splitMulti = false): string[] {
  if (raw === null || raw === undefined) {
    return [];
  }
  if (Array.isArray(raw)) {
    const flat: string[] = [];
    raw.forEach((item) => toValues(item, splitMulti).forEach((v) => flat.push(v)));
    return flat;
  }
  if (typeof raw === 'object') {
    const label = fromObject(raw as ILabelled);
    return label ? [label] : [];
  }
  if (typeof raw === 'boolean') {
    return [raw ? 'Yes' : 'No'];
  }
  return fromString(String(raw), splitMulti);
}

/** Dates are grouped by year: day-level groups are rarely useful as categories. */
export function parseDate(raw: unknown): Date | undefined {
  if (!raw) {
    return undefined;
  }
  const text = String(raw);
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  const year = /\b(19|20)\d{2}\b/.exec(text);
  return year ? new Date(parseInt(year[0], 10), 0, 1) : undefined;
}

export function yearOf(raw: unknown): string[] {
  const date = parseDate(raw);
  return date ? [String(date.getFullYear())] : [];
}

const OFFICE_EXTENSIONS = ['doc', 'docx', 'docm', 'xls', 'xlsx', 'xlsm', 'xlsb', 'ppt', 'pptx', 'pptm', 'vsdx', 'one'];

/** Office files open in the browser; anything else opens (or downloads) directly. */
export function isOfficeFile(fileType: string): boolean {
  return OFFICE_EXTENSIONS.indexOf((fileType || '').toLowerCase()) >= 0;
}

export function encodePath(serverRelativePath: string): string {
  return serverRelativePath
    .split('/')
    .map((segment) => encodeURIComponent(segment))
    .join('/');
}
