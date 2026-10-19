import { IDocumentRecord, IMetadataField, MODIFIED_YEAR_KEY, NO_VALUE } from '../models/IMetadata';

export type FilterState = { [fieldKey: string]: string[] };

export interface IFacetValue {
  value: string;
  count: number;
}

export interface IGroup {
  value: string;
  count: number;
  documents: IDocumentRecord[];
  subgroups?: IGroup[];
}

export interface ICoverage {
  field: IMetadataField;
  filled: number;
  percent: number;
  distinct: number;
}

function valuesOf(doc: IDocumentRecord, fieldKey: string): string[] {
  const values = doc.values[fieldKey];
  return values && values.length ? values : [NO_VALUE];
}

/** AND across fields, OR within one field; the text box matches file names and every value. */
export function filterDocuments(docs: IDocumentRecord[], filters: FilterState, text: string): IDocumentRecord[] {
  const needle = text.trim().toLowerCase();
  const active = Object.keys(filters).filter((k) => filters[k].length);
  return docs.filter((doc) => {
    for (const key of active) {
      const wanted = filters[key];
      if (!valuesOf(doc, key).some((v) => wanted.indexOf(v) >= 0)) {
        return false;
      }
    }
    if (!needle) {
      return true;
    }
    if (doc.name.toLowerCase().indexOf(needle) >= 0) {
      return true;
    }
    return Object.keys(doc.values).some((k) => doc.values[k].some((v) => v.toLowerCase().indexOf(needle) >= 0));
  });
}

/** Years read best in date order (newest first); everything else by size. */
export function isChronological(field: IMetadataField): boolean {
  return field.kind === 'date' || field.key === MODIFIED_YEAR_KEY;
}

function sorter(chronological: boolean): (a: IFacetValue, b: IFacetValue) => number {
  return (a, b) => {
    // "(No value)" always last: it is the gap to fix, not a category.
    if (a.value === NO_VALUE) {
      return 1;
    }
    if (b.value === NO_VALUE) {
      return -1;
    }
    if (chronological) {
      return b.value.localeCompare(a.value);
    }
    return b.count - a.count || a.value.localeCompare(b.value);
  };
}

export function facetValues(docs: IDocumentRecord[], field: IMetadataField): IFacetValue[] {
  const fieldKey = field.key;
  const counts: { [value: string]: number } = {};
  docs.forEach((doc) => {
    valuesOf(doc, fieldKey).forEach((v) => {
      counts[v] = (counts[v] || 0) + 1;
    });
  });
  return Object.keys(counts)
    .map((value) => ({ value, count: counts[value] }))
    .sort(sorter(isChronological(field)));
}

/** A document with several values (multi-value columns) appears in each of its groups. */
export function groupDocuments(docs: IDocumentRecord[], field: IMetadataField, thenBy?: IMetadataField): IGroup[] {
  const fieldKey = field.key;
  const buckets: { [value: string]: IDocumentRecord[] } = {};
  docs.forEach((doc) => {
    valuesOf(doc, fieldKey).forEach((v) => {
      (buckets[v] = buckets[v] || []).push(doc);
    });
  });
  return Object.keys(buckets)
    .map((value) => ({ value, count: buckets[value].length }))
    .sort(sorter(isChronological(field)))
    .map(({ value, count }) => ({
      value,
      count,
      documents: buckets[value],
      subgroups: thenBy && thenBy.key !== fieldKey ? groupDocuments(buckets[value], thenBy) : undefined
    }));
}

/** How many documents have each field filled in: the pilot's metadata-coverage measure. */
export function coverage(docs: IDocumentRecord[], fields: IMetadataField[]): ICoverage[] {
  return fields.map((field) => {
    const distinct: { [v: string]: boolean } = {};
    let filled = 0;
    docs.forEach((doc) => {
      const values = doc.values[field.key];
      if (values && values.length) {
        filled++;
        values.forEach((v) => {
          distinct[v] = true;
        });
      }
    });
    return {
      field,
      filled,
      percent: docs.length ? Math.round((filled / docs.length) * 100) : 0,
      distinct: Object.keys(distinct).length
    };
  });
}

function csvCell(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

export function toCsv(docs: IDocumentRecord[], fields: IMetadataField[]): string {
  const header = ['Name', 'URL', 'Modified'].concat(fields.map((f) => f.title));
  const lines = docs.map((doc) =>
    [doc.name, doc.url, doc.modified ? doc.modified.toISOString() : '']
      .concat(fields.map((f) => (doc.values[f.key] || []).join('; ')))
      .map(csvCell)
      .join(',')
  );
  return [header.map(csvCell).join(',')].concat(lines).join('\r\n');
}
