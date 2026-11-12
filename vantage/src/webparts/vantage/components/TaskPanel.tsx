import * as React from 'react';
import {
  Checkbox, DatePicker, DayOfWeek, DefaultButton, Dropdown, IconButton, MessageBar, MessageBarType, Panel, PanelType,
  Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize, TextField
} from '@fluentui/react';

import styles from './VantageApp.module.scss';
import { Badge, canEdit, errorText, formatDate, type IPersonValue, isOverdue, parseDate, PeoplePicker, relativeTime, toNoonIso, useApp, useConfirm } from './common';
import { type IBoardSummary, type IRelatedRecord, type ITask, PRIORITIES, WORK_STATUSES } from '../models/Models';

interface IForm {
  title: string;
  description: string;
  status: string;
  priority: string;
  assignee: IPersonValue | undefined;
  due: Date | undefined;
}

const toForm = (t?: ITask): IForm => ({
  title: t?.Title || '',
  description: t?.WorkDescription || '',
  status: t?.WorkStatus || 'Not Started',
  priority: t?.Priority || 'Medium',
  assignee: t?.AssignedToKey ? { key: t.AssignedToKey, name: t.AssignedToName || t.AssignedToKey.split('|').pop() as string } : undefined,
  due: t?.DueDate ? new Date(t.DueDate) : undefined
});

export interface ITaskPanelProps {
  /** undefined = closed, 'new' = create, otherwise the task to view/edit. */
  task: ITask | 'new' | undefined;
  board: IBoardSummary | undefined;
  allTasks: ITask[];
  onDismiss: () => void;
  onSaved: () => void;
}

export const TaskPanel: React.FC<ITaskPanelProps> = ({ task, board, allTasks, onDismiss, onSaved }) => {
  const { work, workspace, role, perform, busy, themeStyle } = useApp();
  const existing = task && task !== 'new' ? task : undefined;
  const editable = canEdit(role);
  const [form, setForm] = React.useState<IForm>(toForm(existing));
  const [related, setRelated] = React.useState<IRelatedRecord[] | undefined>(undefined);
  const [comment, setComment] = React.useState('');
  const [checkItem, setCheckItem] = React.useState('');
  const [depTarget, setDepTarget] = React.useState<string>('');
  const [error, setError] = React.useState('');
  const [confirm, confirmDialog] = useConfirm();

  const loadRelated = React.useCallback(async (): Promise<void> => {
    if (!workspace || !existing) return;
    setRelated(await work.listRelated(workspace, existing.Id));
  }, [workspace, existing, work]);

  React.useEffect(() => {
    setForm(toForm(existing));
    setError('');
    setRelated(undefined);
    setComment(''); setCheckItem(''); setDepTarget('');
    loadRelated().catch(e => setError(errorText(e)));
  }, [task]);

  if (!workspace) return null;

  const dirty = JSON.stringify(form) !== JSON.stringify(toForm(existing));
  const save = (): Promise<boolean> => perform(existing ? 'Saving task…' : 'Creating task…', async () => {
    setError('');
    if (!board && !existing) throw new Error('Select a board first.');
    const input = {
      title: form.title, description: form.description, status: form.status, priority: form.priority,
      assignedToKey: form.assignee?.key, assignedToName: form.assignee?.name, dueDate: form.due ? toNoonIso(form.due) : undefined
    };
    if (existing) await work.updateTask(workspace, existing, input);
    else await work.createTask(workspace, (board as IBoardSummary).id, input);
    onSaved();
  });

  /** Never lose typed changes silently (Escape, light dismiss, close button). */
  const requestClose = async (): Promise<void> => {
    if (dirty && editable) {
      const ok = await confirm({ title: 'Discard changes?', confirmText: 'Discard', danger: true, text: 'Your unsaved changes to this task will be lost.' });
      if (!ok) return;
    }
    onDismiss();
  };

  const act = (label: string, fn: () => Promise<void>): void => {
    perform(label, async () => { await fn(); await loadRelated(); }).catch(() => undefined);
  };

  const comments = (related || []).filter(r => r.RecordType === 'Comment');
  const checklist = (related || []).filter(r => r.RecordType === 'Checklist');
  const deps = (related || []).filter(r => r.RecordType === 'Dependency');
  const activity = (related || []).filter(r => r.RecordType === 'Activity').slice().reverse();
  const done = checklist.filter(c => c.ChecklistDone).length;

  const footer = (
    <div className={styles.panelFooter}>
      {editable && <PrimaryButton text={existing ? 'Save changes' : 'Create task'} onClick={() => { save().catch(() => undefined); }}
        disabled={busy || !form.title.trim() || (!!existing && !dirty)} />}
      <DefaultButton text={dirty && editable ? 'Discard' : 'Close'} onClick={() => { requestClose().catch(() => undefined); }} />
    </div>
  );

  const details = (
    <div className={styles.formStack}>
      <TextField label="Title" required value={form.title} maxLength={255} disabled={!editable}
        onChange={(_, v) => setForm({ ...form, title: v || '' })} autoFocus={!existing} />
      <div className={styles.formRow}>
        <Dropdown label="Status" selectedKey={form.status} disabled={!editable}
          options={WORK_STATUSES.map(s => ({ key: s, text: s }))} onChange={(_, o) => setForm({ ...form, status: String(o?.key) })} />
        <Dropdown label="Priority" selectedKey={form.priority} disabled={!editable}
          options={PRIORITIES.map(p => ({ key: p, text: p }))} onChange={(_, o) => setForm({ ...form, priority: String(o?.key) })} />
      </div>
      <div className={styles.formRow}>
        <PeoplePicker label="Assigned to" value={form.assignee} disabled={!editable} onChange={v => setForm({ ...form, assignee: v })} />
        <DatePicker label="Due date" value={form.due} disabled={!editable} firstDayOfWeek={DayOfWeek.Monday} formatDate={formatDate}
          allowTextInput parseDateFromString={parseDate} placeholder="No due date" onSelectDate={d => setForm({ ...form, due: d || undefined })} />
      </div>
      {form.due && editable && <DefaultButton text="Clear due date" onClick={() => setForm({ ...form, due: undefined })} className={styles.linkButton} />}
      <TextField label="Description" multiline autoAdjustHeight rows={4} value={form.description} disabled={!editable}
        onChange={(_, v) => setForm({ ...form, description: v || '' })} />
      {existing && (
        <p className={styles.muted}>
          Task #{existing.Id} · created {relativeTime(existing.Created)} · updated {relativeTime(existing.Modified)}
        </p>
      )}
    </div>
  );

  return (
    <Panel
      isOpen={!!task}
      onDismiss={ev => {
        // Escape pressed inside a picker or dropdown closes that control's popup; it must not close the panel too.
        const target = ev && (ev.target as HTMLElement | null);
        if (target && typeof target.closest === 'function' && target.closest('.ms-DatePicker, .ms-BasePicker, .ms-Dropdown, .ms-Callout')) return;
        requestClose().catch(() => undefined);
      }}
      type={PanelType.medium}
      headerText={existing ? existing.Title : 'New task'}
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => footer}
    >
      <div style={themeStyle}>
      {existing && (
        <div className={styles.panelMeta}>
          <Badge value={existing.WorkStatus} /> <Badge value={existing.Priority} />
          {isOverdue(existing.DueDate, existing.WorkStatus) && <Badge value="Overdue" tone={styles.bad} />}
          {board && <span className={styles.muted}>{board.title}</span>}
        </div>
      )}
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      {!existing ? details : (
        <Pivot aria-label="Task sections" overflowBehavior="menu">
          <PivotItem headerText="Details" itemKey="details"><div className={styles.section}>{details}</div></PivotItem>

          <PivotItem headerText="Comments" itemKey="comments" itemCount={comments.length}>
            <div className={styles.section}>
              {related === undefined ? <Spinner size={SpinnerSize.small} /> : (
                <>
                  {editable && (
                    <div className={styles.formStack}>
                      <TextField label="Add a comment" multiline autoAdjustHeight rows={2} value={comment} onChange={(_, v) => setComment(v || '')} />
                      <PrimaryButton text="Post comment" disabled={busy || !comment.trim()}
                        onClick={() => act('Posting comment…', async () => { await work.addComment(workspace, existing, comment); setComment(''); })} />
                    </div>
                  )}
                  <ul className={styles.feed}>
                    {comments.slice().reverse().map(c => (
                      <li key={c.Id}>
                        <div className={styles.feedHead}><strong>{c.ActorName || 'Someone'}</strong><span className={styles.muted}>{relativeTime(c.Created)}</span></div>
                        <div className={styles.feedBody}>{c.WorkDescription}</div>
                      </li>
                    ))}
                    {!comments.length && <li className={styles.muted}>No comments yet.</li>}
                  </ul>
                </>
              )}
            </div>
          </PivotItem>

          <PivotItem headerText="Checklist" itemKey="checklist" itemCount={checklist.length ? `${done}/${checklist.length}` : undefined}>
            <div className={styles.section}>
              {editable && (
                <form className={styles.inlineForm} onSubmit={e => { e.preventDefault(); if (checkItem.trim()) act('Adding item…', async () => { await work.addChecklistItem(workspace, existing, checkItem); setCheckItem(''); }); }}>
                  <TextField label="New checklist item" value={checkItem} onChange={(_, v) => setCheckItem(v || '')} maxLength={255} />
                  <PrimaryButton type="submit" text="Add" disabled={busy || !checkItem.trim()} />
                </form>
              )}
              <ul className={styles.checklist}>
                {checklist.map(c => (
                  <li key={c.Id}>
                    <Checkbox label={c.Title} checked={!!c.ChecklistDone} disabled={!editable || busy}
                      onChange={(_, v) => {
                        // Optimistic: tick immediately, then persist (reloaded from SharePoint afterwards).
                        setRelated(rows => (rows || []).map(r => r.Id === c.Id ? { ...r, ChecklistDone: !!v } : r));
                        act('Updating…', () => work.setChecklistDone(workspace, existing, c, !!v));
                      }} />
                    {editable && <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel={`Remove ${c.Title}`} disabled={busy}
                      onClick={() => act('Removing…', () => work.removeRelated(workspace, existing, c))} />}
                  </li>
                ))}
                {!checklist.length && <li className={styles.muted}>No checklist items.</li>}
              </ul>
            </div>
          </PivotItem>

          <PivotItem headerText="Dependencies" itemKey="deps" itemCount={deps.length || undefined}>
            <div className={styles.section}>
              {editable && (
                <div className={styles.inlineForm}>
                  <Dropdown label="This task depends on" placeholder="Select a task" selectedKey={depTarget || undefined}
                    options={allTasks.filter(t => t.Id !== existing.Id && !deps.some(d => d.RelatedTaskId === t.Id)).map(t => ({ key: String(t.Id), text: `${t.Title} (#${t.Id})` }))}
                    onChange={(_, o) => setDepTarget(String(o?.key || ''))} />
                  <PrimaryButton text="Add" disabled={busy || !depTarget}
                    onClick={() => act('Adding dependency…', async () => {
                      const target = allTasks.find(t => String(t.Id) === depTarget);
                      if (target) await work.addDependency(workspace, existing, target, deps);
                      setDepTarget('');
                    })} />
                </div>
              )}
              <ul className={styles.checklist}>
                {deps.map(d => {
                  const target = allTasks.find(t => t.Id === d.RelatedTaskId);
                  return (
                    <li key={d.Id}>
                      <span>{target ? target.Title : d.Title} {target && <Badge value={target.WorkStatus} />}</span>
                      {editable && <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove dependency" disabled={busy}
                        onClick={() => act('Removing…', () => work.removeRelated(workspace, existing, d))} />}
                    </li>
                  );
                })}
                {!deps.length && <li className={styles.muted}>No dependencies.</li>}
              </ul>
            </div>
          </PivotItem>

          <PivotItem headerText="Activity" itemKey="activity">
            <div className={styles.section}>
              <ul className={styles.feed}>
                {activity.map(a => (
                  <li key={a.Id}>
                    <div className={styles.feedHead}><strong>{a.ActorName || 'Someone'}</strong><span className={styles.muted}>{relativeTime(a.Created)}</span></div>
                    <div className={styles.feedBody}>{a.Title}</div>
                  </li>
                ))}
                {!activity.length && <li className={styles.muted}>No activity recorded.</li>}
              </ul>
            </div>
          </PivotItem>
        </Pivot>
      )}
      {existing && editable && (
        <div className={styles.dangerZone}>
          <DefaultButton text="Delete task" iconProps={{ iconName: 'Delete' }} disabled={busy}
            onClick={async () => {
              const ok = await confirm({
                title: 'Delete this task?', confirmText: 'Delete', danger: true,
                text: <>“{existing.Title}” and its comments, checklist items, dependencies and activity will be permanently deleted. This cannot be undone. To keep a record, complete the task and let the archive policy move it instead.</>
              });
              if (ok && await perform('Deleting task…', () => work.deleteTask(workspace, existing.Id))) onSaved();
            }} />
        </div>
      )}
      </div>
      {confirmDialog}
    </Panel>
  );
};
