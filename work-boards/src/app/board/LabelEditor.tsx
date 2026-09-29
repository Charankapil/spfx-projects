import * as React from 'react';
import { Dialog, DialogFooter, PrimaryButton, DefaultButton, TextField, Callout, Icon, TooltipHost } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoardColumn, ILabel } from '../../models/types';
import { Swatches } from '../common/Common';
import { shortId } from '../../engine/ids';
import { PALETTE } from '../../models/colors';

export interface ILabelEditorProps {
  column: IBoardColumn;
  onSave: (labels: ILabel[], renames: { from: string; to: string }[]) => Promise<void>;
  onClose: () => void;
}

/** Edit Status labels or Dropdown options: text, colour, order and (for Status) which label means done. */
export function LabelEditor(props: ILabelEditorProps): JSX.Element {
  const original = props.column.labels || [];
  const [labels, setLabels] = React.useState<ILabel[]>(original.map(l => ({ ...l, id: l.id || 'l_' + shortId(5) })));
  const [colorFor, setColorFor] = React.useState<{ index: number; target: HTMLElement } | null>(null);
  const [saving, setSaving] = React.useState(false);
  const isStatus = props.column.type === 'status';

  const update = (index: number, changes: Partial<ILabel>): void => {
    setLabels(ls => ls.map((l, i) => (i === index ? { ...l, ...changes } : changes.isDone && isStatus ? { ...l, isDone: false } : l)));
  };
  const move = (index: number, delta: number): void => {
    setLabels(ls => {
      const next = ls.slice();
      const [l] = next.splice(index, 1);
      next.splice(index + delta, 0, l);
      return next;
    });
  };
  const add = (): void => {
    const used = labels.map(l => l.color);
    const color = (PALETTE.filter(p => used.indexOf(p.value) < 0)[0] || PALETTE[0]).value;
    setLabels(ls => ls.concat([{ id: 'l_' + shortId(5), text: '', color }]));
  };

  const cleaned = labels.map(l => ({ ...l, text: l.text.trim() })).filter(l => l.text.length > 0);
  const texts = cleaned.map(l => l.text.toLowerCase());
  const duplicate = texts.some((t, i) => texts.indexOf(t) !== i);

  const save = async (): Promise<void> => {
    const renames: { from: string; to: string }[] = [];
    cleaned.forEach(l => {
      const before = original.filter(o => o.id === l.id)[0];
      if (before && before.text !== l.text) {
        renames.push({ from: before.text, to: l.text });
      }
    });
    setSaving(true);
    try {
      await props.onSave(cleaned, renames);
    } catch {
      // useBoard shows the error on the board.
      setSaving(false);
      return;
    }
    props.onClose();
  };

  return (
    <Dialog hidden={false} onDismiss={props.onClose} minWidth={460}
      dialogContentProps={{ title: `${isStatus ? 'Labels' : 'Options'} for ${props.column.title}`, subText: 'Renaming a label also updates the items that use it.' }}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {labels.map((l, i) => (
          <div key={l.id} className={styles.labelEditRow}>
            <button type="button" className={styles.swatch} style={{ background: l.color, width: 28, height: 28 }}
              onClick={e => setColorFor({ index: i, target: e.currentTarget })} aria-label={`Colour for ${l.text || 'new label'}`} />
            <TextField value={l.text} onChange={(_, v) => update(i, { text: v || '' })} styles={{ root: { flex: 1 } }}
              ariaLabel="Label text" placeholder="Label" autoFocus={l.text === '' && i === labels.length - 1} />
            {isStatus && (
              <TooltipHost content="Items with this label count as done">
                <button type="button" className={styles.iconBtn} onClick={() => update(i, { isDone: !l.isDone })} aria-pressed={!!l.isDone}
                  aria-label={`Mark ${l.text} as the done label`} style={{ color: l.isDone ? '#00c875' : undefined }}>
                  <Icon iconName={l.isDone ? 'CompletedSolid' : 'Completed'} />
                </button>
              </TooltipHost>
            )}
            <button type="button" className={styles.iconBtn} disabled={i === 0} onClick={() => move(i, -1)} aria-label="Move up"><Icon iconName="Up" /></button>
            <button type="button" className={styles.iconBtn} disabled={i === labels.length - 1} onClick={() => move(i, 1)} aria-label="Move down"><Icon iconName="Down" /></button>
            <button type="button" className={styles.iconBtn} onClick={() => setLabels(ls => ls.filter((_, j) => j !== i))} aria-label={`Remove ${l.text}`}><Icon iconName="Delete" /></button>
          </div>
        ))}
        <button type="button" className={styles.linkBtn} onClick={add}><Icon iconName="Add" /> Add {isStatus ? 'label' : 'option'}</button>
        {duplicate && <span style={{ color: 'var(--wb-danger)' }}>Two labels have the same text. Give each a different name.</span>}
      </div>
      {colorFor && (
        <Callout target={colorFor.target} onDismiss={() => setColorFor(null)} setInitialFocus>
          <div className={styles.popup}>
            <Swatches value={labels[colorFor.index].color} onChange={c => { update(colorFor.index, { color: c }); setColorFor(null); }} />
          </div>
        </Callout>
      )}
      <DialogFooter>
        <PrimaryButton text={saving ? 'Saving…' : 'Save'} disabled={saving || duplicate} onClick={() => { save().catch(() => undefined); }} />
        <DefaultButton text="Cancel" onClick={props.onClose} />
      </DialogFooter>
    </Dialog>
  );
}
