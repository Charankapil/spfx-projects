import * as React from 'react';
import { Checkbox, DefaultButton, Dropdown, Icon, MessageBar, TextField } from '@fluentui/react';

import { FieldType, IField, IRecipient } from '../models/types';
import { shortId } from '../services/hash';
import { IPageBox, PdfPages } from './PdfPages';
import { FIELD_META, recipientColour, tint } from './shared';
import styles from './ESign.module.scss';

export interface IFieldDesignerProps {
  bytes: ArrayBuffer;
  signers: IRecipient[];
  fields: IField[];
  onChange: (fields: IField[]) => void;
  onPageCount?: (count: number) => void;
}

interface IDrag {
  id: string;
  mode: 'move' | 'resize';
  startX: number;
  startY: number;
  origin: IField;
  page: IPageBox;
}

const clamp = (v: number, min: number, max: number): number => Math.max(min, Math.min(max, v));

/**
 * Place fields for each signer: pick a signer and a field type, click on the
 * page to drop it, drag to move, drag the corner to resize. Positions are
 * kept as fractions of the page so they are independent of zoom.
 */
export function FieldDesigner(props: IFieldDesignerProps): JSX.Element {
  const { signers, fields, onChange } = props;
  const [recipientId, setRecipientId] = React.useState(signers[0]?.id);
  const [placing, setPlacing] = React.useState<FieldType>();
  const [selectedId, setSelectedId] = React.useState<string>();
  const drag = React.useRef<IDrag>();

  React.useEffect(() => {
    if (!signers.some(s => s.id === recipientId)) {
      setRecipientId(signers[0]?.id);
    }
  }, [signers, recipientId]);

  const colourOf = (id: string): string => recipientColour(Math.max(0, signers.findIndex(s => s.id === id)));
  const selected = fields.find(f => f.id === selectedId);
  const update = (id: string, change: Partial<IField>): void => onChange(fields.map(f => (f.id === id ? { ...f, ...change } : f)));
  const remove = (id: string): void => {
    onChange(fields.filter(f => f.id !== id));
    setSelectedId(undefined);
  };

  const place = (page: IPageBox, x: number, y: number): void => {
    if (!placing || !recipientId) {
      setSelectedId(undefined);
      return;
    }
    const meta = FIELD_META[placing];
    // Fields keep a constant physical size; height is relative to a portrait page.
    const h = (meta.h * (page.width * 1.294)) / page.height;
    const field: IField = {
      id: shortId(),
      type: placing,
      recipientId,
      page: page.index,
      x: clamp(x / page.width - meta.w / 2, 0, 1 - meta.w),
      y: clamp(y / page.height - h / 2, 0, 1 - h),
      w: meta.w,
      h,
      required: placing !== 'checkbox'
    };
    onChange(fields.concat(field));
    setSelectedId(field.id);
  };

  const onPointerDown = (e: React.PointerEvent, field: IField, page: IPageBox, mode: 'move' | 'resize'): void => {
    e.stopPropagation();
    e.preventDefault();
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
    setSelectedId(field.id);
    drag.current = { id: field.id, mode, startX: e.clientX, startY: e.clientY, origin: field, page };
  };

  const onPointerMove = (e: React.PointerEvent): void => {
    const d = drag.current;
    if (!d) {
      return;
    }
    const dx = (e.clientX - d.startX) / d.page.width;
    const dy = (e.clientY - d.startY) / d.page.height;
    const o = d.origin;
    if (d.mode === 'move') {
      update(d.id, { x: clamp(o.x + dx, 0, 1 - o.w), y: clamp(o.y + dy, 0, 1 - o.h) });
    } else {
      update(d.id, { w: clamp(o.w + dx, 0.02, 1 - o.x), h: clamp(o.h + dy, 0.012, 1 - o.y) });
    }
  };

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (!selected) {
      return;
    }
    const step = e.shiftKey ? 0.02 : 0.004;
    const moves: { [key: string]: Partial<IField> } = {
      ArrowLeft: { x: clamp(selected.x - step, 0, 1 - selected.w) },
      ArrowRight: { x: clamp(selected.x + step, 0, 1 - selected.w) },
      ArrowUp: { y: clamp(selected.y - step, 0, 1 - selected.h) },
      ArrowDown: { y: clamp(selected.y + step, 0, 1 - selected.h) }
    };
    if (moves[e.key]) {
      e.preventDefault();
      update(selected.id, moves[e.key]);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      remove(selected.id);
    }
  };

  const overlay = (page: IPageBox): React.ReactNode =>
    fields
      .filter(f => f.page === page.index)
      .map(f => {
        const colour = colourOf(f.recipientId);
        const signer = signers.find(s => s.id === f.recipientId);
        return (
          <div
            key={f.id}
            tabIndex={0}
            role="button"
            aria-label={`${FIELD_META[f.type].label} for ${signer?.name || 'signer'}`}
            className={`${styles.field} ${styles.fieldDesign} ${f.id === selectedId ? styles.fieldSelected : ''}`}
            style={{
              left: f.x * page.width,
              top: f.y * page.height,
              width: f.w * page.width,
              height: f.h * page.height,
              borderColor: colour,
              background: tint(colour, 0.14),
              color: colour
            }}
            onPointerDown={e => onPointerDown(e, f, page, 'move')}
            onPointerMove={onPointerMove}
            onPointerUp={() => (drag.current = undefined)}
            onKeyDown={onKeyDown}
            onFocus={() => setSelectedId(f.id)}
            title={`${FIELD_META[f.type].label}${f.required ? ' (required)' : ''} - ${signer?.name || ''}`}
          >
            {f.type === 'checkbox' ? '' : f.label || FIELD_META[f.type].label}
            {f.required && f.type !== 'checkbox' ? ' *' : ''}
            {f.id === selectedId && (
              <>
                <button className={styles.remove} aria-label="Remove field" onPointerDown={e => e.stopPropagation()} onClick={() => remove(f.id)}>
                  ×
                </button>
                <span
                  className={styles.resize}
                  onPointerDown={e => onPointerDown(e, f, page, 'resize')}
                  onPointerMove={onPointerMove}
                  onPointerUp={() => (drag.current = undefined)}
                />
              </>
            )}
          </div>
        );
      });

  const missing = signers.filter(s => !fields.some(f => f.recipientId === s.id && f.type === 'signature'));

  return (
    <div className={styles.split}>
      <PdfPages
        bytes={props.bytes}
        onLoaded={props.onPageCount}
        onPageClick={place}
        pageClassName={placing ? styles.placing : ''}
        renderOverlay={overlay}
      />
      <div className={`${styles.card} ${styles.side}`}>
        <h3 className={styles.cardTitle}>Fields</h3>
        <Dropdown
          label="Fields for"
          selectedKey={recipientId}
          onChange={(_, o) => setRecipientId(o?.key as string)}
          options={signers.map((s, i) => ({ key: s.id, text: `${s.name || s.email || `Signer ${i + 1}`}`, data: i }))}
          onRenderOption={o => (
            <span>
              <span className={styles.recipientDot} style={{ background: recipientColour(o?.data as number) }} />
              {o?.text}
            </span>
          )}
        />
        <div className={styles.palette}>
          {(Object.keys(FIELD_META) as FieldType[]).map(type => (
            <DefaultButton
              key={type}
              checked={placing === type}
              toggle
              onClick={() => setPlacing(placing === type ? undefined : type)}
              iconProps={{ iconName: FIELD_META[type].icon }}
              text={FIELD_META[type].label}
              styles={{ root: { padding: '0 6px', minWidth: 0 }, label: { fontWeight: 400, fontSize: 12 } }}
            />
          ))}
        </div>
        <p className={styles.muted}>
          {placing
            ? `Click on the document to place a ${FIELD_META[placing].label.toLowerCase()} field. Click the button again to stop.`
            : 'Choose a field type, then click on the document. Drag fields to move them; drag the dark corner to resize. Arrow keys nudge, Delete removes.'}
        </p>

        {selected && (
          <div className={styles.card} style={{ padding: 12, marginTop: 8 }}>
            <strong>
              <Icon iconName={FIELD_META[selected.type].icon} /> {FIELD_META[selected.type].label}
            </strong>
            <Dropdown
              label="Assigned to"
              selectedKey={selected.recipientId}
              onChange={(_, o) => update(selected.id, { recipientId: o?.key as string })}
              options={signers.map((s, i) => ({ key: s.id, text: s.name || s.email || `Signer ${i + 1}` }))}
            />
            {selected.type === 'text' && (
              <TextField label="Label / hint" value={selected.label || ''} onChange={(_, v) => update(selected.id, { label: v })} />
            )}
            {selected.type !== 'date' && selected.type !== 'name' && selected.type !== 'email' && (
              <Checkbox
                styles={{ root: { marginTop: 8 } }}
                label="Required"
                checked={selected.required}
                onChange={(_, v) => update(selected.id, { required: !!v })}
              />
            )}
            <DefaultButton styles={{ root: { marginTop: 8 } }} iconProps={{ iconName: 'Delete' }} text="Remove" onClick={() => remove(selected.id)} />
          </div>
        )}

        <h4>Placed</h4>
        {signers.map((s, i) => (
          <div key={s.id} className={styles.muted}>
            <span className={styles.recipientDot} style={{ background: recipientColour(i) }} />
            {s.name || s.email}: {fields.filter(f => f.recipientId === s.id).length} field(s)
          </div>
        ))}
        {missing.length > 0 && (
          <MessageBar styles={{ root: { marginTop: 8 } }}>
            Each signer needs at least one signature field. Missing: {missing.map(m => m.name || m.email).join(', ')}
          </MessageBar>
        )}
      </div>
    </div>
  );
}
