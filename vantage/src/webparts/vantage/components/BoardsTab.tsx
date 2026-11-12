import * as React from 'react';
import {
  CommandBar, DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogFooter, DialogType, Dropdown, type IColumn,
  type ICommandBarItemProps, Link, MessageBarType, PrimaryButton, SearchBox, SelectionMode, Shimmer, TextField, Toggle
} from '@fluentui/react';

import styles from './VantageApp.module.scss';
import { Badge, canAdmin, canEdit, Card, EmptyState, formatDate, isOverdue, useApp } from './common';
import { TaskPanel } from './TaskPanel';
import { SharePointService } from '../services/SharePointService';
import { type IBoardSummary, type ITask, TERMINAL_STATUSES, WORK_STATUSES } from '../models/Models';

const PRIORITY_RANK: Record<string, number> = { Critical: 0, High: 1, Medium: 2, Low: 3 };

export const BoardsTab: React.FC = () => {
  const { provisioner, work, workspace, role, perform, busy, notify, user, refreshToken, bumpRefresh } = useApp();
  const [boards, setBoards] = React.useState<IBoardSummary[] | undefined>(undefined);
  const [boardId, setBoardId] = React.useState<string>('');
  const [tasks, setTasks] = React.useState<ITask[] | undefined>(undefined);
  const [openTask, setOpenTask] = React.useState<ITask | 'new' | undefined>(undefined);
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<string>('open');
  const [mineOnly, setMineOnly] = React.useState(false);
  const [newBoard, setNewBoard] = React.useState<{ name: string; description: string } | undefined>(undefined);

  const board = boards && boards.find(b => b.id === boardId);

  React.useEffect(() => {
    if (!workspace) return;
    let live = true;
    setBoards(undefined);
    provisioner.listBoards(workspace.id)
      .then(list => { if (!live) return; setBoards(list); setBoardId(current => list.some(b => b.id === current) ? current : (list[0] ? list[0].id : '')); })
      .catch(e => { if (live) { setBoards([]); notify(MessageBarType.error, `Could not load boards: ${(e as Error).message}`); } });
    return () => { live = false; };
  }, [workspace, provisioner, notify]);

  const loadTasks = React.useCallback(async (): Promise<void> => {
    if (!workspace || !boardId) { setTasks([]); return; }
    setTasks(await work.listTasks(workspace, boardId));
  }, [workspace, boardId, work]);

  React.useEffect(() => {
    setTasks(undefined);
    loadTasks().catch(e => { setTasks([]); notify(MessageBarType.error, `Could not load tasks: ${(e as Error).message}`); });
  }, [loadTasks, refreshToken, notify]);

  if (!workspace) return null;
  if (!boards) return <Card><Shimmer /><Shimmer width="70%" /><Shimmer width="40%" /></Card>;

  const createBoard = async (): Promise<void> => {
    if (!newBoard) return;
    const ok = await perform('Creating board…', async () => {
      const b = await provisioner.createBoard(workspace.id, newBoard.name, newBoard.description, boards);
      setBoards(boards.concat([b]).sort((a, z) => a.title.localeCompare(z.title)));
      setBoardId(b.id);
      notify(MessageBarType.success, `Board “${b.title}” created.`);
    });
    if (ok) setNewBoard(undefined);
  };

  const me = SharePointService.claims(user.loginName).toLowerCase();
  const term = search.trim().toLowerCase();
  const visible = (tasks || [])
    .filter(t => statusFilter === 'all' ? true : statusFilter === 'open' ? TERMINAL_STATUSES.indexOf(t.WorkStatus) < 0 : t.WorkStatus === statusFilter)
    .filter(t => !mineOnly || (t.AssignedToKey || '').toLowerCase() === me)
    .filter(t => !term || t.Title.toLowerCase().indexOf(term) >= 0 || (t.AssignedToName || '').toLowerCase().indexOf(term) >= 0)
    .sort((a, b) => {
      const ad = a.DueDate ? new Date(a.DueDate).getTime() : Infinity;
      const bd = b.DueDate ? new Date(b.DueDate).getTime() : Infinity;
      return ad !== bd ? ad - bd : (PRIORITY_RANK[a.Priority || 'Medium'] - PRIORITY_RANK[b.Priority || 'Medium']) || b.Id - a.Id;
    });

  const columns: IColumn[] = [
    { key: 'title', name: 'Task', minWidth: 200, isRowHeader: true, isResizable: true,
      onRender: (t: ITask) => <Link onClick={() => setOpenTask(t)} className={styles.taskLink}>{t.Title}</Link> },
    { key: 'status', name: 'Status', minWidth: 120, maxWidth: 140, onRender: (t: ITask) => <Badge value={t.WorkStatus} /> },
    { key: 'priority', name: 'Priority', minWidth: 90, maxWidth: 110, onRender: (t: ITask) => <Badge value={t.Priority} /> },
    { key: 'assignee', name: 'Assigned to', minWidth: 120, maxWidth: 200, onRender: (t: ITask) => t.AssignedToName || <span className={styles.muted}>Unassigned</span> },
    { key: 'due', name: 'Due', minWidth: 90, maxWidth: 120,
      onRender: (t: ITask) => t.DueDate ? <span className={isOverdue(t.DueDate, t.WorkStatus) ? styles.bad : undefined}>{formatDate(t.DueDate)}</span> : '—' }
  ];

  const commands: ICommandBarItemProps[] = [
    { key: 'new', text: 'New task', iconProps: { iconName: 'Add' }, disabled: !canEdit(role) || !board || busy, onClick: () => setOpenTask('new') },
    { key: 'refresh', text: 'Refresh', iconProps: { iconName: 'Refresh' }, disabled: busy, onClick: () => { loadTasks().catch(() => undefined); } }
  ];

  return (
    <div className={styles.stack}>
      <div className={styles.toolbar}>
        <Dropdown className={styles.boardPicker} ariaLabel="Board" placeholder={boards.length ? 'Select a board' : 'No boards yet'}
          selectedKey={boardId || undefined} disabled={!boards.length}
          options={boards.map(b => ({ key: b.id, text: b.title }))} onChange={(_, o) => setBoardId(String(o?.key || ''))} />
        {canAdmin(role) && <DefaultButton text="New board" iconProps={{ iconName: 'BoardAdd' }} onClick={() => setNewBoard({ name: '', description: '' })} disabled={busy} />}
        {board && board.description && <span className={styles.muted}>{board.description}</span>}
      </div>

      {!boards.length ? (
        <Card>
          <EmptyState title="Create the first board"
            text={canAdmin(role) ? 'Boards group related tasks, for example by project or team.' : 'A workspace administrator needs to create a board before tasks can be added.'}
            action={canAdmin(role) ? <PrimaryButton text="New board" onClick={() => setNewBoard({ name: '', description: '' })} /> : undefined} />
        </Card>
      ) : (
        <Card>
          <div className={styles.filterBar}>
            <CommandBar items={commands} className={styles.commandBar} ariaLabel="Task actions" />
            <div className={styles.filters}>
              <SearchBox placeholder="Filter tasks" value={search} onChange={(_, v) => setSearch(v || '')} className={styles.search} />
              <Dropdown ariaLabel="Status filter" selectedKey={statusFilter} className={styles.statusFilter}
                options={[{ key: 'open', text: 'Open tasks' }, { key: 'all', text: 'All tasks' }].concat(WORK_STATUSES.map(s => ({ key: s, text: s })))}
                onChange={(_, o) => setStatusFilter(String(o?.key))} />
              <Toggle inlineLabel label="Assigned to me" checked={mineOnly} onChange={(_, v) => setMineOnly(!!v)} />
            </div>
          </div>
          {tasks === undefined ? <><Shimmer /><Shimmer /><Shimmer width="80%" /></> : visible.length ? (
            <>
              <div className={styles.wideOnly}>
                <DetailsList items={visible} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified}
                  getKey={(t: ITask) => String(t.Id)} onItemInvoked={(t: ITask) => setOpenTask(t)} />
              </div>
              <ul className={styles.narrowOnly} aria-label="Tasks">
                {visible.map(t => (
                  <li key={t.Id} className={styles.taskCard}>
                    <button type="button" className={styles.taskCardButton} onClick={() => setOpenTask(t)}>
                      <span className={styles.taskCardTitle}>{t.Title}</span>
                      <span className={styles.taskCardMeta}>
                        <Badge value={t.WorkStatus} /> <Badge value={t.Priority} />
                        {t.DueDate && <span className={isOverdue(t.DueDate, t.WorkStatus) ? styles.bad : styles.muted}>Due {formatDate(t.DueDate)}</span>}
                      </span>
                      <span className={styles.muted}>{t.AssignedToName || 'Unassigned'}</span>
                    </button>
                  </li>
                ))}
              </ul>
            </>
          ) : (
            <EmptyState title={tasks.length ? 'No tasks match the filters' : 'No tasks on this board yet'}
              action={!tasks.length && canEdit(role) ? <PrimaryButton text="New task" onClick={() => setOpenTask('new')} /> : undefined} />
          )}
          {tasks && tasks.length >= 500 && <p className={styles.muted}>Showing the 500 most recent tasks on this board.</p>}
        </Card>
      )}

      <TaskPanel task={openTask} board={board} allTasks={tasks || []}
        onDismiss={() => setOpenTask(undefined)}
        onSaved={() => { setOpenTask(undefined); loadTasks().catch(() => undefined); bumpRefresh(); }} />

      <Dialog hidden={!newBoard} onDismiss={() => setNewBoard(undefined)} minWidth={420}
        dialogContentProps={{ type: DialogType.normal, title: 'New board', subText: 'Boards share the workspace’s permissions. Use a separate workspace for work that needs different access.' }}>
        <div className={styles.formStack}>
          <TextField label="Board name" required maxLength={120} value={newBoard?.name || ''} autoFocus
            onChange={(_, v) => setNewBoard({ name: v || '', description: newBoard ? newBoard.description : '' })} />
          <TextField label="Description" multiline rows={2} value={newBoard?.description || ''}
            onChange={(_, v) => setNewBoard({ name: newBoard ? newBoard.name : '', description: v || '' })} />
        </div>
        <DialogFooter>
          <PrimaryButton text="Create board" disabled={busy || !newBoard || !newBoard.name.trim()} onClick={() => { createBoard().catch(() => undefined); }} />
          <DefaultButton text="Cancel" onClick={() => setNewBoard(undefined)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
