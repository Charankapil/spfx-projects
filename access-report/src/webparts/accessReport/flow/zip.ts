/**
 * A minimal ZIP writer (stored entries, no compression) - enough to hand Power Automate a
 * package built in the browser, with nothing to install. UTF-8 names, CRC-32 checked.
 */

let table: number[] | undefined;

export function crc32(bytes: Uint8Array): number {
  if (!table) {
    table = [];
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table.push(c >>> 0);
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = table[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export function utf8(text: string): Uint8Array {
  // encodeURIComponent yields UTF-8 percent escapes; works on every browser SharePoint supports.
  const escaped = encodeURIComponent(text);
  const out: number[] = [];
  for (let i = 0; i < escaped.length; i++) {
    const ch = escaped.charAt(i);
    if (ch === '%') {
      out.push(parseInt(escaped.substring(i + 1, i + 3), 16));
      i += 2;
    } else {
      out.push(ch.charCodeAt(0));
    }
  }
  return new Uint8Array(out);
}

export interface IZipEntry {
  /** Path inside the archive, with forward slashes. */
  name: string;
  text: string;
}

export function zip(entries: IZipEntry[]): Uint8Array {
  const chunks: number[][] = [];
  const central: number[][] = [];
  let offset = 0;
  const u16 = (n: number): number[] => [n & 0xff, (n >>> 8) & 0xff];
  const u32 = (n: number): number[] => [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff];
  // A fixed timestamp (2026-01-01 00:00) keeps the package identical from build to build.
  const time = u16(0);
  const date = u16(((2026 - 1980) << 9) | (1 << 5) | 1);
  entries.forEach((e) => {
    const name = utf8(e.name);
    const data = utf8(e.text);
    const crc = u32(crc32(data));
    const size = u32(data.length);
    // version 2.0, method 0 = stored; flag bit 11 (UTF-8 names) only when a name needs it,
    // so plain names give the most conventional archive possible.
    const flags = name.length === e.name.length ? 0 : 0x0800;
    const common = ([] as number[]).concat(u16(20), u16(flags), u16(0), time, date, crc, size, size, u16(name.length), u16(0));
    const local = ([] as number[]).concat(u32(0x04034b50), common, Array.prototype.slice.call(name), Array.prototype.slice.call(data));
    central.push(([] as number[]).concat(u32(0x02014b50), u16(20), common, u16(0), u16(0), u16(0), u32(0), u32(offset), Array.prototype.slice.call(name)));
    chunks.push(local);
    offset += local.length;
  });
  const centralSize = central.reduce((n, c) => n + c.length, 0);
  const end = ([] as number[]).concat(u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length), u32(centralSize), u32(offset), u16(0));
  const out = new Uint8Array(offset + centralSize + end.length);
  let at = 0;
  chunks.concat(central, [end]).forEach((c) => {
    out.set(c, at);
    at += c.length;
  });
  return out;
}
