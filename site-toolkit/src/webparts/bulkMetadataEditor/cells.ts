import { fromDateInput, toDateInput } from '../../common/format';

export type FieldKind = 'Text' | 'Number' | 'Currency' | 'Choice' | 'Boolean' | 'DateTime';

export interface IColumn {
  name: string;
  title: string;
  kind: FieldKind;
  choices: string[];
}

/** The text shown in a cell for a stored value. */
export function display(value: unknown, kind: FieldKind): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (kind === 'Boolean') {
    return value ? 'Yes' : 'No';
  }
  if (kind === 'DateTime') {
    return toDateInput(String(value));
  }
  return String(value);
}

export interface IParsed {
  ok: boolean;
  value?: unknown;
}

/** Turns cell text into the value SharePoint stores; rejects text that does not fit the column. */
export function parse(text: string, column: IColumn): IParsed {
  const trimmed = text.trim();
  if (trimmed === '') {
    // eslint-disable-next-line @rushstack/no-new-null
    return column.kind === 'Text' ? { ok: true, value: '' } : { ok: true, value: null };
  }
  switch (column.kind) {
    case 'Number':
    case 'Currency': {
      const n = Number(trimmed);
      return isNaN(n) ? { ok: false } : { ok: true, value: n };
    }
    case 'Boolean': {
      if (/^(yes|true|1)$/i.test(trimmed)) {
        return { ok: true, value: true };
      }
      return /^(no|false|0)$/i.test(trimmed) ? { ok: true, value: false } : { ok: false };
    }
    case 'DateTime': {
      const iso = fromDateInput(trimmed);
      return iso ? { ok: true, value: iso } : { ok: false };
    }
    case 'Choice': {
      const match = column.choices.filter((c) => c.toLowerCase() === trimmed.toLowerCase())[0];
      return match ? { ok: true, value: match } : { ok: false };
    }
    default:
      return trimmed.length > 255 ? { ok: false } : { ok: true, value: trimmed };
  }
}
