import * as React from 'react';
import { FieldType, RequestState } from '../models/types';
import styles from './ESign.module.scss';

export const FIELD_META: { [type in FieldType]: { label: string; icon: string; w: number; h: number } } = {
  signature: { label: 'Signature', icon: 'InsertSignatureLine', w: 0.26, h: 0.06 },
  initials: { label: 'Initials', icon: 'Font', w: 0.1, h: 0.05 },
  date: { label: 'Date signed', icon: 'Calendar', w: 0.17, h: 0.03 },
  name: { label: 'Full name', icon: 'Contact', w: 0.24, h: 0.03 },
  email: { label: 'Email', icon: 'Mail', w: 0.27, h: 0.03 },
  text: { label: 'Text', icon: 'TextField', w: 0.26, h: 0.03 },
  checkbox: { label: 'Checkbox', icon: 'CheckboxComposite', w: 0.03, h: 0.023 }
};

/** One colour per recipient, used for their fields in the designer and signing view. */
const RECIPIENT_COLOURS = ['#0f6cbd', '#c239b3', '#107c10', '#ca5010', '#5c2e91', '#038387', '#a4262c', '#8764b8', '#498205', '#004e8c'];

export function recipientColour(index: number): string {
  return RECIPIENT_COLOURS[index % RECIPIENT_COLOURS.length];
}

/** Translucent fill for a recipient's colour. */
export function tint(hex: string, alpha: number): string {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

const STATE_LABELS: { [state in RequestState]: string } = {
  InProgress: 'In progress',
  Completed: 'Completed',
  Declined: 'Declined',
  Cancelled: 'Cancelled',
  Expired: 'Expired'
};

export function StatusBadge(props: { state: RequestState; actionRequired?: boolean }): JSX.Element {
  if (props.actionRequired) {
    return <span className={`${styles.badge} ${styles.badgeAction}`}>Action required</span>;
  }
  return <span className={`${styles.badge} ${styles[`badge${props.state}`]}`}>{STATE_LABELS[props.state]}</span>;
}

export function formatDate(iso: string | undefined, withTime = false): string {
  if (!iso) {
    return '-';
  }
  const date = new Date(iso);
  return withTime
    ? date.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' } as Intl.DateTimeFormatOptions)
    : date.toLocaleDateString(undefined, { dateStyle: 'medium' } as Intl.DateTimeFormatOptions);
}

export function downloadBlob(data: BlobPart, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
