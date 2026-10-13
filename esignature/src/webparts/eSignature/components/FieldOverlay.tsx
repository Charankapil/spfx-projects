import * as React from 'react';
import { IField, IRequest, ISignedPayload } from '../models/types';
import { IRequestStatus } from '../services/status';
import { IPageBox } from './PdfPages';
import styles from './ESign.module.scss';

/** Positioning style for a field on a rendered page. */
export function fieldBox(field: IField, page: IPageBox): React.CSSProperties {
  return {
    left: field.x * page.width,
    top: field.y * page.height,
    width: field.w * page.width,
    height: field.h * page.height
  };
}

/** Renders a filled-in value the way it will appear in the signed PDF. */
export function FieldValue(props: { field: IField; value: string; page: IPageBox }): JSX.Element | null {
  const { field, value } = props;
  if (!value) {
    return null;
  }
  if (field.type === 'signature' || field.type === 'initials') {
    return <img src={value} alt={field.type} className={styles.fieldImage} />;
  }
  if (field.type === 'checkbox') {
    return value === 'true' ? <span style={{ fontSize: Math.max(10, props.page.height * field.h * 0.8) }}>✓</span> : null;
  }
  return (
    <span style={{ fontSize: Math.max(8, Math.min(14, props.page.height * field.h * 0.6)), color: '#0d0d33', whiteSpace: 'nowrap' }}>{value}</span>
  );
}

/** Read-only overlay of everything signed so far - what the document looks like right now. */
export function signedValuesOverlay(request: IRequest, status: IRequestStatus, page: IPageBox, exceptSlot?: string): React.ReactNode {
  const out: React.ReactNode[] = [];
  for (const slot of status.slots) {
    if (!slot.signedActivity || slot.original.id === exceptSlot) {
      continue;
    }
    const values = (slot.signedActivity.payload as ISignedPayload).values || {};
    request.payload.fields
      .filter(f => f.recipientId === slot.original.id && f.page === page.index && values[f.id])
      .forEach(f =>
        out.push(
          <div key={f.id} className={`${styles.field} ${styles.fieldReadonly}`} style={fieldBox(f, page)}>
            <FieldValue field={f} value={values[f.id]} page={page} />
          </div>
        )
      );
  }
  return out;
}
