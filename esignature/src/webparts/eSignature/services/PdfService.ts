import type { PDFDocument, PDFFont, PDFPage } from 'pdf-lib';
import type { IPdfDocument } from 'pdfjs-dist/legacy/build/pdf';

import { IActivity, IField, IRequest, ISignedPayload } from '../models/types';
import { IRequestStatus } from './status';

/**
 * PDF work happens in the browser: pdf.js renders pages for the designer and
 * signing screens, pdf-lib writes the final signed PDF. Both are loaded on
 * demand so the dashboard stays light.
 */

type PdfJs = typeof import('pdfjs-dist/legacy/build/pdf');
type PdfLib = typeof import('pdf-lib');

let pdfJs: Promise<PdfJs> | undefined;
let pdfLib: Promise<PdfLib> | undefined;

export function loadPdfJs(): Promise<PdfJs> {
  if (!pdfJs) {
    // The worker entry registers pdf.js's worker code on window, so pdf.js
    // runs it on the main thread instead of fetching a worker script from a
    // CDN (which tenant content security policies commonly block).
    pdfJs = import(/* webpackChunkName: 'esign-pdfjs-worker' */ 'pdfjs-dist/legacy/build/pdf.worker.entry').then(() =>
      import(/* webpackChunkName: 'esign-pdfjs' */ 'pdfjs-dist/legacy/build/pdf')
    );
  }
  return pdfJs;
}

export function loadPdfLib(): Promise<PdfLib> {
  if (!pdfLib) {
    pdfLib = import(/* webpackChunkName: 'esign-pdflib' */ 'pdf-lib');
  }
  return pdfLib;
}

export async function openPdf(bytes: ArrayBuffer): Promise<IPdfDocument> {
  const pdfjs = await loadPdfJs();
  // pdf.js takes ownership of (detaches) the buffer it is given, so hand it a copy.
  return pdfjs.getDocument({ data: new Uint8Array(bytes.slice(0)), isEvalSupported: false }).promise;
}

export function dataUrlToBytes(dataUrl: string): Uint8Array {
  const base64 = dataUrl.substring(dataUrl.indexOf(',') + 1);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** Standard PDF fonts only cover WinAnsi; swap anything else for '?' rather than failing. */
function safe(font: PDFFont, value: string): string {
  const supported = font.getCharacterSet();
  let out = '';
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    out += supported.indexOf(code) >= 0 ? value[i] : code === 10 ? ' ' : '?';
  }
  return out;
}

function fitFontSize(font: PDFFont, text: string, width: number, height: number): number {
  let size = Math.min(height * 0.7, 14);
  while (size > 5 && font.widthOfTextAtSize(text, size) > width - 4) {
    size -= 0.5;
  }
  return size;
}

export interface IStampOptions {
  request: IRequest;
  status: IRequestStatus;
  activities: IActivity[];
  completedOn: Date;
}

/**
 * Builds the signed PDF: every signer's values drawn into their fields, then
 * a certificate of completion appended with the hashes and the full event
 * log. The returned bytes are what gets saved and hashed as the final copy.
 */
export async function buildSignedPdf(original: ArrayBuffer, options: IStampOptions): Promise<Uint8Array> {
  const lib = await loadPdfLib();
  const doc = await lib.PDFDocument.load(original, { updateMetadata: false });
  const font = await doc.embedFont(lib.StandardFonts.Helvetica);
  const bold = await doc.embedFont(lib.StandardFonts.HelveticaBold);
  const pages = doc.getPages();
  const { request, status } = options;

  for (const slot of status.slots) {
    if (!slot.signedActivity) {
      continue;
    }
    const values = (slot.signedActivity.payload as ISignedPayload).values || {};
    for (const field of request.payload.fields.filter(f => f.recipientId === slot.original.id)) {
      const value = values[field.id];
      const page = pages[field.page];
      if (!value || !page) {
        continue;
      }
      await drawField(lib, doc, page, field, value, font);
    }
  }

  await appendCertificate(lib, doc, font, bold, options);

  doc.setTitle(request.payload.title);
  doc.setSubject(`Signed via SharePoint eSignature - request ${request.requestId}`);
  doc.setKeywords(['eSignature', request.requestId]);
  doc.setProducer('SharePoint eSignature (SPFx)');
  doc.setModificationDate(options.completedOn);
  return doc.save();
}

async function drawField(lib: PdfLib, doc: PDFDocument, page: PDFPage, field: IField, value: string, font: PDFFont): Promise<void> {
  const { width, height } = page.getSize();
  // Fields are stored top-left based as fractions of the page; PDF space is bottom-left in points.
  const x = field.x * width;
  const w = field.w * width;
  const h = field.h * height;
  const y = height - field.y * height - h;

  if (field.type === 'signature' || field.type === 'initials') {
    if (value.indexOf('data:image/png') !== 0) {
      return;
    }
    const image = await doc.embedPng(dataUrlToBytes(value));
    const scale = Math.min(w / image.width, h / image.height);
    const iw = image.width * scale;
    const ih = image.height * scale;
    page.drawImage(image, { x: x + (w - iw) / 2, y: y + (h - ih) / 2, width: iw, height: ih });
    return;
  }
  if (field.type === 'checkbox') {
    if (value === 'true') {
      const inset = Math.min(w, h) * 0.2;
      page.drawLine({ start: { x: x + inset, y: y + h / 2 }, end: { x: x + w * 0.42, y: y + inset }, thickness: 1.6, color: lib.rgb(0, 0, 0) });
      page.drawLine({ start: { x: x + w * 0.42, y: y + inset }, end: { x: x + w - inset, y: y + h - inset }, thickness: 1.6, color: lib.rgb(0, 0, 0) });
    }
    return;
  }
  const text = safe(font, value);
  const size = fitFontSize(font, text, w, h);
  page.drawText(text, { x: x + 2, y: y + (h - size) / 2 + size * 0.2, size, font, color: lib.rgb(0.05, 0.05, 0.2) });
}

const A4 = { w: 595.28, h: 841.89 };
const MARGIN = 48;

class Writer {
  public page!: PDFPage;
  public y = 0;

  constructor(private lib: PdfLib, private doc: PDFDocument, private font: PDFFont, private bold: PDFFont) {
    this.newPage();
  }

  public newPage(): void {
    this.page = this.doc.addPage([A4.w, A4.h]);
    this.y = A4.h - MARGIN;
  }

  public ensure(space: number): void {
    if (this.y - space < MARGIN) {
      this.newPage();
    }
  }

  public text(value: string, options: { size?: number; bold?: boolean; indent?: number; color?: number } = {}): void {
    const size = options.size || 9;
    const font = options.bold ? this.bold : this.font;
    const indent = options.indent || 0;
    const maxWidth = A4.w - MARGIN * 2 - indent;
    for (const line of this.wrap(safe(font, value), font, size, maxWidth)) {
      this.ensure(size + 4);
      this.page.drawText(line, {
        x: MARGIN + indent,
        y: this.y - size,
        size,
        font,
        color: this.lib.rgb(options.color || 0.1, options.color || 0.1, options.color || 0.1)
      });
      this.y -= size + 4;
    }
  }

  public row(label: string, value: string): void {
    this.ensure(14);
    const top = this.y;
    this.page.drawText(safe(this.bold, label), { x: MARGIN, y: top - 9, size: 9, font: this.bold, color: this.lib.rgb(0.35, 0.35, 0.35) });
    this.text(value, { indent: 130 });
    if (this.y > top - 13) {
      this.y = top - 13;
    }
  }

  public rule(): void {
    this.ensure(10);
    this.page.drawLine({
      start: { x: MARGIN, y: this.y - 4 },
      end: { x: A4.w - MARGIN, y: this.y - 4 },
      thickness: 0.5,
      color: this.lib.rgb(0.8, 0.8, 0.8)
    });
    this.y -= 10;
  }

  public gap(points: number): void {
    this.y -= points;
  }

  private wrap(value: string, font: PDFFont, size: number, maxWidth: number): string[] {
    const lines: string[] = [];
    let line = '';
    for (const word of value.split(' ')) {
      const candidate = line ? `${line} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
        line = candidate;
        continue;
      }
      if (line) {
        lines.push(line);
      }
      // Hard-break words longer than a line (hashes, URLs).
      let rest = word;
      while (font.widthOfTextAtSize(rest, size) > maxWidth) {
        let cut = rest.length - 1;
        while (cut > 1 && font.widthOfTextAtSize(rest.substring(0, cut), size) > maxWidth) {
          cut--;
        }
        lines.push(rest.substring(0, cut));
        rest = rest.substring(cut);
      }
      line = rest;
    }
    if (line) {
      lines.push(line);
    }
    return lines.length ? lines : [''];
  }
}

function stamp(iso: string | undefined): string {
  return iso ? `${new Date(iso).toUTCString().replace('GMT', 'UTC')}` : '-';
}

export const EVENT_LABELS: { [type: string]: string } = {
  sent: 'Sent for signature',
  viewed: 'Viewed the document',
  signed: 'Signed',
  declined: 'Declined to sign',
  delegated: 'Delegated signing',
  reminded: 'Sent a reminder',
  cancelled: 'Cancelled the request',
  completed: 'Completed - signed document generated',
  copied: 'Saved a copy next to the original'
};

async function appendCertificate(lib: PdfLib, doc: PDFDocument, font: PDFFont, bold: PDFFont, options: IStampOptions): Promise<void> {
  const { request, status, activities, completedOn } = options;
  const p = request.payload;
  const w = new Writer(lib, doc, font, bold);

  w.text('Certificate of completion', { size: 18, bold: true });
  w.text('Electronic signature audit trail', { size: 10, color: 0.4 });
  w.gap(8);
  w.rule();
  w.row('Request', p.title);
  w.row('Request ID', request.requestId);
  w.row('Document', `${p.documentName} (${p.pageCount} page${p.pageCount === 1 ? '' : 's'})`);
  w.row('Sender', `${p.senderName} <${p.senderEmail}>`);
  w.row('Sent', stamp(p.sentOn));
  w.row('Completed', stamp(completedOn.toISOString()));
  w.row('Original SHA-256', p.documentHash);
  w.row('Signing order', status.slots.some(s => s.original.order !== status.slots[0].original.order) ? 'Sequential / mixed' : 'Parallel');
  w.gap(6);
  w.text(
    'The SHA-256 fingerprint above identifies the exact document that was sent. Each signer below recorded the fingerprint of the document they ' +
      'were shown; a match proves they signed this content. The fingerprint of this signed file is recorded in the SharePoint activity log ' +
      'and can be checked with Verify in the eSignature web part.',
    { size: 8, color: 0.4 }
  );
  w.gap(10);

  w.text('Signers', { size: 13, bold: true });
  w.rule();
  for (const slot of status.slots) {
    const signedBy = slot.signedActivity;
    const signed = signedBy ? (signedBy.payload as ISignedPayload) : undefined;
    w.ensure(110);
    w.text(`${slot.current.name} <${slot.current.email}>`, { size: 10, bold: true });
    if (slot.delegations.length) {
      w.row('Delegated from', `${slot.original.name} <${slot.original.email}>`);
    }
    w.row('Signing order', String(slot.original.order));
    w.row('Viewed', stamp(slot.viewedOn));
    w.row('Signed', stamp(slot.signedOn));
    if (signed && signedBy) {
      w.row('Identity', `SharePoint account ${signedBy.authorName} <${signedBy.authorEmail}>`);
      w.row('Document SHA-256', signed.documentHash === p.documentHash ? `${signed.documentHash} (matches)` : `${signed.documentHash} (DOES NOT MATCH)`);
      w.row('Time zone', signed.timeZone || '-');
      w.row('Browser', signed.userAgent || '-');
      w.row('Consent', signed.consent || '-');
      const sig = request.payload.fields.find(f => f.recipientId === slot.original.id && f.type === 'signature');
      const value = sig ? signed.values[sig.id] : undefined;
      if (value && value.indexOf('data:image/png') === 0) {
        const image = await doc.embedPng(dataUrlToBytes(value));
        const scale = Math.min(160 / image.width, 48 / image.height);
        w.ensure(56);
        w.page.drawImage(image, { x: MARGIN + 130, y: w.y - image.height * scale, width: image.width * scale, height: image.height * scale });
        w.gap(image.height * scale + 6);
      }
    }
    w.rule();
  }
  if (status.ccs.length) {
    w.row('Copied (CC)', status.ccs.map(c => `${c.name} <${c.email}>`).join(', '));
    w.gap(6);
  }

  w.gap(6);
  w.text('Event log', { size: 13, bold: true });
  w.rule();
  const ordered = activities
    .filter(a => a.type !== 'completed' && a.type !== 'copied')
    .sort((a, b) => new Date(a.created).getTime() - new Date(b.created).getTime());
  for (const a of ordered) {
    const note = status.ignored[a.id] ? `  -  NOT COUNTED: ${status.ignored[a.id]}` : '';
    w.text(`${stamp(a.created)}  -  ${EVENT_LABELS[a.type] || a.type}  -  ${a.authorName} <${a.authorEmail}>${note}`, { size: 8 });
  }
  w.text(`${stamp(completedOn.toISOString())}  -  ${EVENT_LABELS.completed}`, { size: 8 });
  if (status.warnings.length) {
    w.gap(8);
    w.text('Integrity warnings', { size: 11, bold: true });
    status.warnings.forEach(warning => w.text(warning, { size: 8 }));
  }
}

export async function pageSizes(bytes: ArrayBuffer): Promise<{ count: number }> {
  const doc = await openPdf(bytes);
  const count = doc.numPages;
  await doc.destroy();
  return { count };
}
