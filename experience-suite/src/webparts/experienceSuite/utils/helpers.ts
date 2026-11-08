import { IListItem } from '../models/Models';

/** Case-insensitive, whitespace-tolerant e-mail comparison. Empty values never match. */
export function sameEmail(a: string | undefined, b: string | undefined): boolean {
  const left = (a || '').trim().toLowerCase();
  const right = (b || '').trim().toLowerCase();
  return left.length > 0 && left === right;
}

/** The e-mail an approval is assigned to: ApproverEmail, falling back to Owner. */
export function approvalAssignee(item: IListItem): string {
  return (item.ApproverEmail || item.Owner || '').toString().trim();
}

export function isPending(item: IListItem): boolean {
  return (item.Status || '').trim().toLowerCase() === 'pending';
}

/** Turns a status such as "In Progress" into a single CSS-safe class name ("in-progress"). */
export function statusClass(status: string | undefined): string {
  return (status || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

/** Only allows http(s) and mailto links through; anything else (e.g. javascript:) becomes undefined. */
export function safeUrl(url: string | undefined): string | undefined {
  const value = (url || '').trim();
  return /^(https?:|mailto:)/i.test(value) ? value : undefined;
}

/** Combines free text typed by the user with a fixed KQL restriction. */
export function buildKql(userText: string, restriction: string): string {
  const clean = (userText || '').trim();
  return clean ? `(${clean}) AND ${restriction}` : restriction;
}

/** Formats an ISO date for display; returns '' for missing or invalid values. */
export function formatDate(value: string | undefined): string {
  if (!value) return '';
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date.toLocaleDateString();
}
