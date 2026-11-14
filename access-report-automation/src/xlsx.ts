import { deflateRawSync } from 'node:zlib';

/**
 * A small Excel (.xlsx) writer with no dependency: several sheets, a bold frozen header row,
 * column widths, a filter on the header, text and numbers. Text is written as inline strings,
 * so nothing a site owner typed into a title can ever be read as a formula.
 */

export type Cell = string | number | null | undefined;

export interface Sheet {
  name: string;
  /** Column widths in characters; also the number of columns. */
  widths: number[];
  /** First row is the header when `header` is true. */
  rows: Cell[][];
  header?: boolean;
  /** Rows (0-based) to draw bold besides the header, e.g. section titles on a summary sheet. */
  bold?: number[];
}

/** Excel's limits. */
export const MAX_ROWS = 1048576;
const MAX_CELL = 32767;

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

/** A ZIP archive (deflate) from named buffers. */
export function zip(files: Array<{ name: string; data: Buffer }>): Buffer {
  const parts: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8');
    const crc = crc32(f.data);
    const packed = deflateRawSync(f.data, { level: 6 });
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0x0800, 6); // UTF-8 names
    local.writeUInt16LE(8, 8); // deflate
    local.writeUInt16LE(0, 10); // time
    local.writeUInt16LE(0x21, 12); // date: 1980-01-01
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(f.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    parts.push(local, name, packed);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0);
    c.writeUInt16LE(20, 4);
    c.writeUInt16LE(20, 6);
    c.writeUInt16LE(0x0800, 8);
    c.writeUInt16LE(8, 10);
    c.writeUInt16LE(0, 12);
    c.writeUInt16LE(0x21, 14);
    c.writeUInt32LE(crc, 16);
    c.writeUInt32LE(packed.length, 20);
    c.writeUInt32LE(f.data.length, 24);
    c.writeUInt16LE(name.length, 28);
    c.writeUInt32LE(offset, 42);
    central.push(c, name);
    offset += local.length + name.length + packed.length;
  }
  const centralSize = central.reduce((n, b) => n + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat(parts.concat(central, [end]));
}

/** Text that is safe inside XML: escaped, and without the characters XML 1.0 forbids. */
export function xmlText(v: string): string {
  let s = v.length > MAX_CELL ? v.substring(0, MAX_CELL - 1) + '…' : v;
  // eslint-disable-next-line no-control-regex
  s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '').replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '');
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function columnName(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Sheet names: at most 31 characters, none of \ / ? * [ ] : and unique. */
export function sheetNames(names: string[]): string[] {
  const used = new Set<string>();
  return names.map((raw) => {
    let base = raw.replace(/[\\/?*[\]:]/g, ' ').replace(/^'+|'+$/g, '').trim().substring(0, 31) || 'Sheet';
    let name = base;
    for (let i = 2; used.has(name.toLowerCase()); i++) {
      const suffix = ` (${i})`;
      base = base.substring(0, 31 - suffix.length);
      name = base + suffix;
    }
    used.add(name.toLowerCase());
    return name;
  });
}

function sheetXml(s: Sheet): Buffer {
  const chunks: string[] = [];
  const cols = s.widths.length;
  const rows = s.rows.slice(0, MAX_ROWS);
  const lastRef = `${columnName(Math.max(0, cols - 1))}${Math.max(1, rows.length)}`;
  chunks.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
  chunks.push(`<dimension ref="A1:${lastRef}"/>`);
  chunks.push('<sheetViews><sheetView workbookViewId="0">');
  if (s.header) {
    chunks.push('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>');
  }
  chunks.push('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
  chunks.push('<cols>' + s.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>');
  chunks.push('<sheetData>');
  const bold = new Set(s.bold || []);
  let part: string[] = [];
  rows.forEach((row, r) => {
    const style = (s.header && r === 0) || bold.has(r) ? ' s="1"' : '';
    part.push(`<row r="${r + 1}">`);
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined || v === '') {
        continue;
      }
      const ref = `${columnName(c)}${r + 1}`;
      if (typeof v === 'number' && isFinite(v)) {
        part.push(`<c r="${ref}"${style}><v>${v}</v></c>`);
      } else {
        part.push(`<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlText(String(v))}</t></is></c>`);
      }
    }
    part.push('</row>');
    if (part.length > 20000) {
      chunks.push(part.join(''));
      part = [];
    }
  });
  chunks.push(part.join(''));
  chunks.push('</sheetData>');
  if (s.header && rows.length > 1 && cols > 0) {
    chunks.push(`<autoFilter ref="A1:${lastRef}"/>`);
  }
  chunks.push('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
  return Buffer.concat(chunks.map((c) => Buffer.from(c, 'utf8')));
}

export function buildXlsx(sheets: Sheet[]): Buffer {
  if (!sheets.length) {
    throw new Error('A workbook needs at least one sheet.');
  }
  const names = sheetNames(sheets.map((s) => s.name));
  const x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const files: Array<{ name: string; data: Buffer }> = [
    {
      name: '[Content_Types].xml',
      data: Buffer.from(
        x +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
          '</Types>'
      )
    },
    {
      name: '_rels/.rels',
      data: Buffer.from(
        x +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
      )
    },
    {
      name: 'xl/workbook.xml',
      data: Buffer.from(
        x +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>' +
          names.map((n, i) => `<sheet name="${xmlText(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
          '</sheets></workbook>'
      )
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: Buffer.from(
        x +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
          `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
          '</Relationships>'
      )
    },
    {
      name: 'xl/styles.xml',
      data: Buffer.from(
        x +
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          '</styleSheet>'
      )
    }
  ];
  sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s) }));
  return zip(files);
}
