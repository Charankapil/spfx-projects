import * as React from 'react';
import { useApp } from '../AppContext';
import { IBoard, IWorkItem, CellValue, ColumnType, IBoardGroup, ILabel, IBoardColumn, IBoardConfig } from '../../models/types';
import { IBoardRights, boardRights, READ_ONLY } from '../../services/permissions';
import { keyForIndex, keyBetween, rebalance, ORDER_STEP } from '../../engine/ordering';
import { shortId } from '../../engine/ids';
import { PALETTE } from '../../models/colors';
import { mapLimit, SpError } from '../../services/SpClient';
import { IListState } from '../../services/ItemService';

const POLL_MS = 20000;

export interface IBoardState {
  board: IBoard | null;
  items: IWorkItem[];
  updateCounts: { [itemId: number]: number };
  rights: IBoardRights;
  loading: boolean;
  error: string;
  notice: string;
}

export interface IBoardActions {
  reload: (quiet?: boolean) => Promise<void>;
  clearMessages: () => void;
  setError: (message: string) => void;
  updateCell: (item: IWorkItem, columnId: string, value: CellValue) => void;
  renameItem: (item: IWorkItem, title: string) => void;
  addItem: (groupId: string, title: string, opts?: { parentId?: number; values?: { [columnId: string]: CellValue }; atTop?: boolean }) => Promise<IWorkItem | null>;
  moveItem: (itemId: number, groupId: string, beforeItemId: number | null) => void;
  moveItemsToGroup: (ids: number[], groupId: string) => Promise<void>;
  deleteItems: (ids: number[]) => Promise<void>;
  addGroup: () => Promise<void>;
  updateGroup: (groupId: string, changes: Partial<IBoardGroup>) => Promise<void>;
  moveGroup: (groupId: string, delta: number) => Promise<void>;
  deleteGroup: (groupId: string) => Promise<void>;
  addColumn: (type: ColumnType, title: string, afterColumnId?: string) => Promise<void>;
  renameColumn: (columnId: string, title: string) => Promise<void>;
  deleteColumn: (columnId: string) => Promise<void>;
  moveColumn: (columnId: string, toIndex: number) => Promise<void>;
  saveLabels: (columnId: string, labels: ILabel[], renames: { from: string; to: string }[]) => Promise<void>;
  addOption: (column: IBoardColumn, text: string) => Promise<void>;
  updateConfig: (mutate: (cfg: IBoardConfig) => IBoardConfig) => Promise<void>;
  setBoard: (board: IBoard) => void;
  bumpUpdateCount: (itemId: number, delta: number) => void;
  applyItem: (item: IWorkItem) => void;
}

function errorText(e: unknown): string {
  if (e instanceof SpError && e.isAccessDenied) {
    return 'You do not have permission to do that on this board.';
  }
  return (e as Error).message || 'Something went wrong.';
}

export function useBoard(boardId: number): [IBoardState, IBoardActions] {
  const app = useApp();
  const { services, me } = app;
  const [state, setState] = React.useState<IBoardState>({
    board: null, items: [], updateCounts: {}, rights: READ_ONLY, loading: true, error: '', notice: ''
  });
  const stateRef = React.useRef(state);
  stateRef.current = state;
  const pending = React.useRef(0);
  const listState = React.useRef<IListState | null>(null);
  const rememberListState = (ls: IListState): void => {
    listState.current = ls;
  };

  const patch = (p: Partial<IBoardState>): void => setState(s => ({ ...s, ...p }));
  const setError = (message: string): void => patch({ error: message });

  const setBoard = (board: IBoard): void => {
    patch({ board });
    app.replaceBoard(board);
  };

  const reload = React.useCallback(async (quiet?: boolean): Promise<void> => {
    if (!quiet) {
      patch({ loading: true, error: '' });
    }
    try {
      const board = await services.boards.getBoard(boardId);
      const [items, perms, counts, ls] = await Promise.all([
        services.items.loadItems(board),
        services.items.listPermissions(board),
        services.updates.counts(board).catch(() => ({})),
        services.items.listState(board)
      ]);
      rememberListState(ls);
      const isOwner = board.ownerIds.indexOf(me.id) >= 0;
      setState(s => ({ ...s, board, items, updateCounts: counts, rights: boardRights(perms, isOwner), loading: false }));
    } catch (e) {
      if (e instanceof SpError && e.isNotFound) {
        patch({ loading: false, board: null, error: 'This board was deleted or you no longer have access to it.' });
      } else {
        patch({ loading: false, error: errorText(e) });
      }
    }
  }, [boardId]);

  React.useEffect(() => {
    setState({ board: null, items: [], updateCounts: {}, rights: READ_ONLY, loading: true, error: '', notice: '' });
    reload().catch(() => undefined);
  }, [boardId]);

  // Near-live refresh: poll the list's last-modified stamp while the tab is visible.
  React.useEffect(() => {
    const timer = setInterval(async () => {
      const b = stateRef.current.board;
      if (!b || document.visibilityState !== 'visible' || pending.current > 0) {
        return;
      }
      try {
        const ls = await services.items.listState(b);
        const prev = listState.current;
        if (prev && (ls.lastModified !== prev.lastModified || ls.lastDeleted !== prev.lastDeleted) && pending.current === 0) {
          await reload(true);
        }
        rememberListState(ls);
      } catch {
        // Ignore polling errors; the next user action will surface real problems.
      }
    }, POLL_MS);
    return () => clearInterval(timer);
  }, [boardId]);

  const track = async <T>(fn: () => Promise<T>): Promise<T> => {
    pending.current++;
    try {
      return await fn();
    } finally {
      pending.current--;
    }
  };

  const replaceItem = (item: IWorkItem): void => {
    setState(s => ({ ...s, items: s.items.map(i => (i.id === item.id ? item : i)) }));
  };

  const saveItem = (item: IWorkItem, next: IWorkItem, changes: Parameters<typeof services.items.updateItem>[2]): void => {
    const board = stateRef.current.board;
    if (!board) {
      return;
    }
    replaceItem(next);
    track(async () => {
      try {
        const res = await services.items.updateItem(board, item, changes);
        if (res.conflicted) {
          const fresh = await services.items.getItem(board, item.id);
          replaceItem(fresh);
          patch({ notice: `Someone else changed "${fresh.title}" at the same time. Both changes were kept.` });
        } else {
          setState(s => ({ ...s, items: s.items.map(i => (i.id === item.id ? { ...i, etag: res.etag || i.etag } : i)) }));
        }
      } catch (e) {
        replaceItem(item);
        setError(errorText(e));
      }
    }).catch(() => undefined);
  };

  const updateCell = (item: IWorkItem, columnId: string, value: CellValue): void => {
    const current = stateRef.current.items.filter(i => i.id === item.id)[0] || item;
    saveItem(current, { ...current, values: { ...current.values, [columnId]: value } }, { values: { [columnId]: value } });
  };

  const renameItem = (item: IWorkItem, title: string): void => {
    const current = stateRef.current.items.filter(i => i.id === item.id)[0] || item;
    if (title === current.title) {
      return;
    }
    saveItem(current, { ...current, title }, { title });
  };

  const siblingsOf = (groupId: string, parentId: number | null, excludeId?: number): IWorkItem[] =>
    stateRef.current.items
      .filter(i => (parentId ? i.parentId === parentId : !i.parentId && i.groupId === groupId) && i.id !== excludeId)
      .sort((a, b) => a.sortOrder - b.sortOrder || a.id - b.id);

  const addItem: IBoardActions['addItem'] = async (groupId, title, opts = {}) => {
    const board = stateRef.current.board;
    if (!board) {
      return null;
    }
    const siblings = siblingsOf(groupId, opts.parentId || null);
    const sortOrder = siblings.length === 0 ? ORDER_STEP
      : opts.atTop ? keyBetween(null, siblings[0].sortOrder) : keyBetween(siblings[siblings.length - 1].sortOrder, null);
    return track(async () => {
      try {
        const created = await services.items.createItem(board, { title, groupId, sortOrder, parentId: opts.parentId || null, values: opts.values });
        setState(s => ({ ...s, items: s.items.concat([created]) }));
        return created;
      } catch (e) {
        setError(errorText(e));
        return null;
      }
    });
  };

  const moveItem: IBoardActions['moveItem'] = (itemId, groupId, beforeItemId) => {
    const item = stateRef.current.items.filter(i => i.id === itemId)[0];
    if (!item) {
      return;
    }
    const siblings = siblingsOf(groupId, item.parentId, item.id);
    let index = beforeItemId === null ? siblings.length : siblings.map(s => s.id).indexOf(beforeItemId);
    if (index < 0) {
      index = siblings.length;
    }
    const { key, rebalance: needs } = keyForIndex(siblings, index);
    if (needs) {
      const ordered = siblings.slice();
      ordered.splice(index, 0, item);
      const changes = rebalance(ordered);
      if (!changes.some(c => c.id === itemId)) {
        // Its key didn't change, but its group may have.
        changes.push({ id: itemId, sortOrder: (index + 1) * ORDER_STEP });
      }
      const board = stateRef.current.board as IBoard;
      setState(s => ({
        ...s,
        items: s.items.map(i => {
          const c = changes.filter(x => x.id === i.id)[0];
          return c ? { ...i, sortOrder: c.sortOrder, groupId: i.id === itemId ? groupId : i.groupId } : i;
        })
      }));
      track(() => mapLimit(changes, 4, c => {
        const it = stateRef.current.items.filter(i => i.id === c.id)[0];
        return services.items.updateItem(board, it, { sortOrder: c.sortOrder, groupId: c.id === itemId ? groupId : undefined });
      })).then(() => reload(true)).catch(e => setError(errorText(e)));
      return;
    }
    // Subitems follow their parent's group.
    const subs = stateRef.current.items.filter(i => i.parentId === itemId);
    saveItem(item, { ...item, groupId, sortOrder: key }, { groupId, sortOrder: key });
    if (!item.parentId && groupId !== item.groupId) {
      subs.forEach(sub => saveItem(sub, { ...sub, groupId }, { groupId }));
    }
  };

  const moveItemsToGroup = async (ids: number[], groupId: string): Promise<void> => {
    for (const id of ids) {
      moveItem(id, groupId, null);
    }
  };

  const deleteItems = async (ids: number[]): Promise<void> => {
    const board = stateRef.current.board;
    if (!board) {
      return;
    }
    const all = stateRef.current.items.filter(i => ids.indexOf(i.id) >= 0 || (i.parentId !== null && ids.indexOf(i.parentId) >= 0)).map(i => i.id);
    const before = stateRef.current.items;
    setState(s => ({ ...s, items: s.items.filter(i => all.indexOf(i.id) < 0) }));
    await track(async () => {
      try {
        await mapLimit(all, 4, id => services.items.recycleItem(board, id));
        patch({ notice: all.length === 1 ? 'Item moved to the site recycle bin.' : `${all.length} items moved to the site recycle bin.` });
      } catch (e) {
        setState(s => ({ ...s, items: before }));
        setError(errorText(e));
        await reload(true);
      }
    });
  };

  const withBoard = async (fn: (b: IBoard) => Promise<IBoard>): Promise<void> => {
    const board = stateRef.current.board;
    if (!board) {
      return;
    }
    await track(async () => {
      try {
        setBoard(await fn(board));
      } catch (e) {
        setError(errorText(e));
      }
    });
  };

  const updateConfig = (mutate: (cfg: IBoardConfig) => IBoardConfig): Promise<void> =>
    withBoard(b => services.boards.updateConfig(b, mutate));

  const addGroup = (): Promise<void> => updateConfig(cfg => {
    const used = cfg.groups.map(g => g.color);
    const color = (PALETTE.filter(p => used.indexOf(p.value) < 0)[0] || PALETTE[0]).value;
    cfg.groups.unshift({ id: 'g_' + shortId(6), title: 'New group', color });
    return cfg;
  });

  const updateGroup = (groupId: string, changes: Partial<IBoardGroup>): Promise<void> => updateConfig(cfg => {
    cfg.groups = cfg.groups.map(g => (g.id === groupId ? { ...g, ...changes, id: g.id } : g));
    return cfg;
  });

  const moveGroup = (groupId: string, delta: number): Promise<void> => updateConfig(cfg => {
    const i = cfg.groups.map(g => g.id).indexOf(groupId);
    const j = i + delta;
    if (i < 0 || j < 0 || j >= cfg.groups.length) {
      return cfg;
    }
    const [g] = cfg.groups.splice(i, 1);
    cfg.groups.splice(j, 0, g);
    return cfg;
  });

  const deleteGroup = async (groupId: string): Promise<void> => {
    const ids = stateRef.current.items.filter(i => i.groupId === groupId && !i.parentId).map(i => i.id);
    if (ids.length > 0) {
      await deleteItems(ids);
    }
    await updateConfig(cfg => {
      if (cfg.groups.length > 1) {
        cfg.groups = cfg.groups.filter(g => g.id !== groupId);
      }
      return cfg;
    });
  };

  const addColumn = async (type: ColumnType, title: string, afterColumnId?: string): Promise<void> => {
    await withBoard(b => services.boards.addColumn(b, type, title, afterColumnId));
    await reload(true);
  };

  const renameColumn = (columnId: string, title: string): Promise<void> => withBoard(b => services.boards.renameColumn(b, columnId, title));

  const deleteColumn = (columnId: string): Promise<void> => withBoard(b => services.boards.deleteColumn(b, columnId));

  const moveColumn = (columnId: string, toIndex: number): Promise<void> => withBoard(b => services.boards.moveColumn(b, columnId, toIndex));

  /** Save labels; items using a renamed label are updated to the new text. */
  const saveLabels = async (columnId: string, labels: ILabel[], renames: { from: string; to: string }[]): Promise<void> => {
    await withBoard(b => services.boards.saveLabels(b, columnId, labels));
    const board = stateRef.current.board;
    const col = board ? board.config.columns.filter(c => c.id === columnId)[0] : null;
    if (!board || !col || renames.length === 0) {
      return;
    }
    const affected: { item: IWorkItem; value: CellValue }[] = [];
    stateRef.current.items.forEach(item => {
      const v = item.values[columnId];
      if (col.type === 'dropdown' && Array.isArray(v)) {
        const next = (v as string[]).map(x => { const r = renames.filter(rn => rn.from === x)[0]; return r ? r.to : x; });
        if (next.join('\n') !== (v as string[]).join('\n')) {
          affected.push({ item, value: next });
        }
      } else if (typeof v === 'string') {
        const r = renames.filter(rn => rn.from === v)[0];
        if (r) {
          affected.push({ item, value: r.to });
        }
      }
    });
    await track(async () => {
      try {
        await mapLimit(affected, 4, a => services.items.updateItem(board, a.item, { values: { [columnId]: a.value } }));
      } catch (e) {
        setError(errorText(e));
      }
      await reload(true);
    });
  };

  const addOption = async (column: IBoardColumn, text: string): Promise<void> => {
    if ((column.labels || []).some(l => l.text === text)) {
      return;
    }
    const used = (column.labels || []).map(l => l.color);
    const color = (PALETTE.filter(p => used.indexOf(p.value) < 0)[0] || PALETTE[0]).value;
    await updateConfig(cfg => {
      cfg.columns.forEach(c => {
        if (c.id === column.id) {
          c.labels = (c.labels || []).concat([{ id: 'o_' + shortId(5), text, color }]);
        }
      });
      return cfg;
    });
  };

  const bumpUpdateCount = (itemId: number, delta: number): void => {
    setState(s => ({ ...s, updateCounts: { ...s.updateCounts, [itemId]: Math.max(0, (s.updateCounts[itemId] || 0) + delta) } }));
  };

  const actions: IBoardActions = {
    reload,
    clearMessages: () => patch({ error: '', notice: '' }),
    setError,
    updateCell,
    renameItem,
    addItem,
    moveItem,
    moveItemsToGroup,
    deleteItems,
    addGroup,
    updateGroup,
    moveGroup,
    deleteGroup,
    addColumn,
    renameColumn,
    deleteColumn,
    moveColumn,
    saveLabels,
    addOption,
    updateConfig,
    setBoard,
    bumpUpdateCount,
    applyItem: replaceItem
  };
  return [state, actions];
}
