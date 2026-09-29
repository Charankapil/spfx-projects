import * as React from 'react';
import { Dialog, DialogFooter, PrimaryButton, DefaultButton, TextField, ComboBox, ChoiceGroup, Checkbox, MessageBar, MessageBarType, Icon, Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { TEMPLATES, IBoardTemplate } from '../../models/templates';
import { IBoard, IPerson, CellValue, ITimelineValue } from '../../models/types';
import { suggestBoardKey, isValidBoardKey } from '../../engine/ids';
import { PALETTE } from '../../models/colors';
import { PeoplePicker, Swatches } from '../common/Common';
import { addDays, todayIso } from '../../engine/dates';
import { ORDER_STEP } from '../../engine/ordering';
import { ItemService } from '../../services/ItemService';

/** Template values use day offsets ("+3") for dates; turn them into real dates. */
function resolveValue(raw: CellValue | string, type: string, today: string): CellValue {
  const day = (v: unknown): string | null => (typeof v === 'string' && /^[+-]\d+$/.test(v) ? addDays(today, parseInt(v, 10)) : (v as string) || null);
  if (type === 'date') {
    return day(raw);
  }
  if (type === 'timeline') {
    const tv = raw as unknown as ITimelineValue;
    return { start: day(tv.start), end: day(tv.end) };
  }
  return raw as CellValue;
}

async function addExampleItems(items: ItemService, board: IBoard, template: IBoardTemplate, me: IPerson): Promise<void> {
  const today = todayIso();
  const ownerCol = board.config.columns.filter(c => c.id === 'owner')[0];
  let order = ORDER_STEP;
  for (const t of template.items) {
    const values: { [id: string]: CellValue } = {};
    Object.keys(t.values).forEach(colId => {
      const col = board.config.columns.filter(c => c.id === colId)[0];
      if (col) {
        values[colId] = resolveValue(t.values[colId], col.type, today);
      }
    });
    if (ownerCol) {
      values[ownerCol.id] = [me];
    }
    const created = await items.createItem(board, { title: t.title, groupId: t.group, sortOrder: order, values });
    order += ORDER_STEP;
    let subOrder = ORDER_STEP;
    for (const s of t.subitems || []) {
      await items.createItem(board, { title: s, groupId: t.group, sortOrder: subOrder, parentId: created.id });
      subOrder += ORDER_STEP;
    }
  }
}

export function NewBoardDialog(props: { onClose: () => void; onCreated: (board: IBoard) => void }): JSX.Element {
  const app = useApp();
  const [templateId, setTemplateId] = React.useState('project');
  const [title, setTitle] = React.useState('');
  const [key, setKey] = React.useState('');
  const [keyTouched, setKeyTouched] = React.useState(false);
  const [folder, setFolder] = React.useState('');
  const [color, setColor] = React.useState(PALETTE[6].value);
  const [description, setDescription] = React.useState('');
  const [privacy, setPrivacy] = React.useState<'Main' | 'Private'>('Main');
  const [members, setMembers] = React.useState<IPerson[]>([]);
  const [examples, setExamples] = React.useState(true);
  const [busy, setBusy] = React.useState('');
  const [error, setError] = React.useState('');

  const takenKeys = app.boards.map(b => b.key);
  const folders = app.boards.map(b => b.folder).filter((f, i, a) => f && a.indexOf(f) === i);
  const template = TEMPLATES.filter(t => t.id === templateId)[0];
  const effectiveKey = keyTouched ? key : suggestBoardKey(title || template.title, takenKeys);
  const keyTaken = takenKeys.some(k => k.toUpperCase() === effectiveKey.toUpperCase());
  const keyValid = isValidBoardKey(effectiveKey);

  const create = async (): Promise<void> => {
    setError('');
    try {
      const board = await app.services.boards.createBoard({
        title: title.trim(), key: effectiveKey, description: description.trim(), folder: folder.trim(), color, privacy,
        templateId, memberIds: members.map(m => m.id)
      }, app.me.id, step => setBusy(step + '…'));
      if (examples && template.items.length > 0) {
        setBusy('Adding example items…');
        await addExampleItems(app.services.items, board, template, app.me);
      }
      setBusy('');
      props.onCreated(board);
    } catch (e) {
      setBusy('');
      setError((e as Error).message);
    }
  };

  return (
    <Dialog hidden={false} onDismiss={busy ? undefined : props.onClose} minWidth={640} maxWidth={760}
      dialogContentProps={{ title: 'New board' }} modalProps={{ isBlocking: !!busy }}>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div>
          <label className={styles.small}>Start from</label>
          <div className={styles.templateGrid} role="radiogroup" aria-label="Template">
            {TEMPLATES.map(t => (
              <button key={t.id} type="button" role="radio" aria-checked={t.id === templateId}
                className={`${styles.templateCard} ${t.id === templateId ? styles.templateCardOn : ''}`} onClick={() => setTemplateId(t.id)}>
                <strong><Icon iconName={t.icon} /> {t.title}</strong>
                <span className={styles.muted}>{t.description}</span>
              </button>
            ))}
          </div>
        </div>
        <TextField label="Board name" required value={title} onChange={(_, v) => setTitle(v || '')} placeholder={template.title} autoFocus />
        <div className={styles.row} style={{ alignItems: 'flex-start' }}>
          <TextField label="Board key" value={effectiveKey} styles={{ root: { width: 160 } }}
            onChange={(_, v) => { setKeyTouched(true); setKey((v || '').toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 8)); }}
            errorMessage={keyTaken ? 'Already used on this site' : !keyValid ? '2–8 letters or digits, starting with a letter' : undefined}
            description={`Item IDs look like ${effectiveKey || 'KEY'}-12`} />
          <ComboBox label="Folder (optional)" allowFreeform autoComplete="on" text={folder} styles={{ root: { width: 220 } }}
            options={folders.map(f => ({ key: f, text: f }))} placeholder="No folder"
            onChange={(_, option, __, value) => setFolder(option ? String(option.text) : value || '')} />
        </div>
        <TextField label="Description (optional)" value={description} onChange={(_, v) => setDescription(v || '')} />
        <div>
          <label className={styles.small}>Colour</label>
          <Swatches value={color} onChange={setColor} />
        </div>
        <ChoiceGroup label="Who can see it" selectedKey={privacy} onChange={(_, o) => o && setPrivacy(o.key as 'Main' | 'Private')}
          options={[
            { key: 'Main', text: 'Everyone on this site (main board)' },
            { key: 'Private', text: 'Only people I choose (private board)', disabled: !app.isSiteOwner }
          ]} />
        {!app.isSiteOwner && <span className={`${styles.small} ${styles.muted}`}>Only site owners can create private boards, because it changes SharePoint permissions.</span>}
        {privacy === 'Private' && <PeoplePicker label="Members" selected={members} onChange={setMembers} />}
        {template.items.length > 0 && <Checkbox label="Add a few example items" checked={examples} onChange={(_, c) => setExamples(!!c)} />}
      </div>
      <DialogFooter>
        {busy && <Spinner size={SpinnerSize.small} label={busy} labelPosition="right" styles={{ root: { display: 'inline-flex', marginRight: 12 } }} />}
        <PrimaryButton text="Create board" disabled={!!busy || !title.trim() || keyTaken || !keyValid} onClick={() => { create().catch(() => undefined); }} />
        <DefaultButton text="Cancel" disabled={!!busy} onClick={props.onClose} />
      </DialogFooter>
    </Dialog>
  );
}
