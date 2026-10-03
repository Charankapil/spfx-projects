/**
 * Small RFC 4180 CSV reader: quoted fields, doubled quotes, line breaks inside
 * quotes, CRLF / LF, a UTF-8 byte-order mark, and comma / semicolon / tab
 * delimiters (detected from the header line, as Excel and Power Automate vary).
 */
function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.substring(1) : text;
}

export function detectDelimiter(text: string): string {
  const firstLine = stripBom(text).split(/\r?\n/, 1)[0] || '';
  const counts = [',', ';', '\t', '|'].map((d) => ({ d, n: firstLine.split(d).length - 1 }));
  counts.sort((a, b) => b.n - a.n);
  return counts[0].n > 0 ? counts[0].d : ',';
}

export function parseCsv(text: string, delimiter?: string): string[][] {
  const src = stripBom(text);
  const d = delimiter || detectDelimiter(src);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < src.length; i++) {
    const c = src.charAt(i);
    if (inQuotes) {
      if (c === '"') {
        if (src.charAt(i + 1) === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"' && field === '') {
      inQuotes = true;
    } else if (c === d) {
      row.push(field);
      field = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && src.charAt(i + 1) === '\n') {
        i++;
      }
      row.push(field);
      field = '';
      if (row.length > 1 || row[0] !== '') {
        rows.push(row);
      }
      row = [];
    } else {
      field += c;
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    if (row.length > 1 || row[0] !== '') {
      rows.push(row);
    }
  }
  return rows;
}
