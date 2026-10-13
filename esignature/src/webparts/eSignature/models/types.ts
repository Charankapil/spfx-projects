/**
 * Shared shapes for eSignature requests.
 *
 * A request is owned by its sender (one item in the Requests list that only
 * the sender can edit). Everything a signer does is a separate item in the
 * Activity list that only that signer can edit, stamped by SharePoint with
 * the real author. The request's state is never stored: it is derived from
 * those two sources every time (see status.ts), so nobody can mark a request
 * signed on someone else's behalf.
 */

export type FieldType = 'signature' | 'initials' | 'date' | 'name' | 'email' | 'text' | 'checkbox';

export type RecipientRole = 'signer' | 'cc';

export interface IRecipient {
  /** Stable slot id - fields and signing actions refer to this, not the email, so delegation works. */
  id: string;
  name: string;
  email: string;
  role: RecipientRole;
  /** Signing order. Equal numbers sign in parallel; lower numbers sign first. */
  order: number;
  /** Optional note added only to this recipient's email (stored with the request, so not secret). */
  privateMessage?: string;
  /** Role label from a template, e.g. "Employee" or "Manager". */
  roleName?: string;
}

export interface IField {
  id: string;
  type: FieldType;
  recipientId: string;
  /** Zero-based page index. */
  page: number;
  /** Position and size as fractions (0..1) of the page width / height, top-left origin. */
  x: number;
  y: number;
  w: number;
  h: number;
  required: boolean;
  label?: string;
}

export interface IRequestPayload {
  /** Schema version of this JSON. */
  v: 1;
  title: string;
  message: string;
  documentName: string;
  /** Server-relative URL of the frozen PDF that is being signed. */
  documentUrl: string;
  /** Server-relative URL of the file the request was started from (may be a Word file). */
  sourceUrl?: string;
  /** SHA-256 (hex) of the frozen PDF at the time of sending. */
  documentHash: string;
  pageCount: number;
  recipients: IRecipient[];
  fields: IField[];
  /** ISO date; the request expires at the end of this day. */
  expiresOn?: string;
  /** Automatic reminder interval in days; 0 = none. */
  reminderDays: number;
  /** Copy the signed PDF next to the source file once completed. */
  saveToSource: boolean;
  senderName: string;
  senderEmail: string;
  sentOn: string;
}

export type RequestState = 'InProgress' | 'Completed' | 'Declined' | 'Cancelled' | 'Expired';

export type ActivityType =
  | 'sent'
  | 'viewed'
  | 'signed'
  | 'declined'
  | 'delegated'
  | 'reminded'
  | 'cancelled'
  | 'completed'
  | 'copied';

export interface ISignedPayload {
  values: { [fieldId: string]: string };
  /** SHA-256 of the frozen PDF the signer saw - proves what they signed. */
  documentHash: string;
  /** SHA-256 of the request payload JSON the signer saw - proves the fields/recipients they agreed to. */
  requestHash: string;
  consent: string;
  userAgent: string;
  timeZone: string;
}

export interface IDelegatedPayload {
  toName: string;
  toEmail: string;
  reason: string;
}

export interface ICompletedPayload {
  signedUrl: string;
  finalHash: string;
}

export interface IActivity {
  id: number;
  requestId: string;
  type: ActivityType;
  /** Recipient slot the activity is about (signed/declined/delegated/viewed). */
  slot?: string;
  /** Who SharePoint says created the item - the only identity that counts. */
  authorEmail: string;
  authorName: string;
  created: string;
  modified: string;
  docHash?: string;
  payload: unknown;
}

export interface IRequest {
  itemId: number;
  requestId: string;
  /** Sender-controlled flag: 'Active' or 'Cancelled'. */
  senderStatus: string;
  payload: IRequestPayload;
  /** Raw JSON as stored, used for the request hash. */
  payloadJson: string;
  authorEmail: string;
  authorName: string;
  created: string;
}

export interface ITemplate {
  itemId?: number;
  name: string;
  description: string;
  /** Roles, not people: name/email are placeholders filled in at send time. */
  recipients: IRecipient[];
  fields: IField[];
  message: string;
  reminderDays: number;
  expiresInDays: number;
  /** Page count of the document the template was designed on. */
  pageCount: number;
}
