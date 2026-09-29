import * as React from 'react';
import { Panel, PanelType, Pivot, PivotItem, PrimaryButton, DefaultButton, Icon, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoard, IBoardColumn, IWorkItem, IUpdate, IAttachment, IActivityEntry, IPerson } from '../../models/types';
import { IBoardRights } from '../../services/permissions';
import { IBoardActions } from '../board/useBoard';
import { useApp } from '../AppContext';
import { Cell } from '../cells/Cells';
import { Avatar, useConfirm } from '../common/Common';
import { isDone } from '../../engine/viewQuery';
import { parseSegments, mentionToken } from '../../engine/mentions';
import { timeAgo } from '../../engine/dates';
import { initials, personColor } from '../../services/PeopleService';

export interface IItemPanelProps {
  board: IBoard;
  item: IWorkItem;
  items: IWorkItem[];
  rights: IBoardRights;
  actions: IBoardActions;
  onClose: () => void;
  onOpenItem: (itemId: number) => void;
  onEditLabels: (column: IBoardColumn) => void;
}

export function ItemPanel(props: IItemPanelProps): JSX.Element {
  const { board, item, rights, actions } = props;
  const { services } = useApp();
  const [title, setTitle] = React.useState(item.title);
  const [tab, setTab] = React.useState('updates');
  React.useEffect(() => setTitle(item.title), [item.id, item.title]);
  const group = board.config.groups.filter(g => g.id === item.groupId)[0] || board.config.groups[0];
  const parent = item.parentId ? props.items.filter(i => i.id === item.parentId)[0] : null;
  const subitems = props.items.filter(i => i.parentId === item.id).sort((a, b) => a.sortOrder - b.sortOrder);
  const done = isDone(board, item);
  const [newSub, setNewSub] = React.useState('');

  const commitTitle = (): void => {
    const t = title.trim();
    if (t && t !== item.title) {
      actions.renameItem(item, t);
    } else {
      setTitle(item.title);
    }
  };

  return (
    <Panel
      isOpen
      type={PanelType.custom}
      customWidth="640px"
      onDismiss={props.onClose}
      isLightDismiss
      closeButtonAriaLabel="Close"
      headerText={`${board.key}-${item.id}`}
      styles={{ headerText: { fontSize: 13, color: 'var(--wb-subtle)', fontFamily: 'monospace' }, content: { paddingBottom: 40 } }}
    >
      <div className={styles.panelHead}>
        {parent && (
          <button type="button" className={styles.linkBtn} onClick={() => props.onOpenItem(parent.id)}>
            <Icon iconName="Back" /> Subitem of {parent.title}
          </button>
        )}
        <input className={styles.panelTitle} value={title} disabled={!rights.canEdit} aria-label="Item name"
          onChange={e => setTitle(e.target.value)} onBlur={commitTitle}
          onKeyDown={e => { if (e.key === 'Enter') { (e.target as HTMLInputElement).blur(); } }} />
        <div className={styles.row}>
          <span className={styles.boardDot} style={{ background: group.color }} />
          <span className={styles.muted}>{group.title}</span>
          <span className={styles.muted}>· Created {timeAgo(item.created)}{item.author ? ' by ' + item.author.title : ''}</span>
        </div>

        <div className={styles.fieldGrid}>
          {board.config.columns.map(col => (
            <React.Fragment key={col.id}>
              <div className={styles.fieldName}>{col.title}</div>
              <div className={styles.fieldValue} style={col.type === 'status' ? { padding: '4px 0' } : undefined}>
                <Cell column={col} value={item.values[col.id]} canEdit={rights.canEdit} canManage={rights.canManage} accent={group.color}
                  done={done} alignLeft onChange={v => actions.updateCell(item, col.id, v)} onEditLabels={props.onEditLabels} onAddOption={actions.addOption} />
              </div>
            </React.Fragment>
          ))}
        </div>

        {!item.parentId && (
          <div>
            <h3 className={styles.sectionTitle} style={{ marginBottom: 6 }}>Subitems {subitems.length > 0 && <span className={styles.muted}>({subitems.length})</span>}</h3>
            {subitems.map(s => (
              <div key={s.id} className={styles.memberRow}>
                <Icon iconName={isDone(board, s) ? 'CompletedSolid' : 'CircleRing'} style={{ color: isDone(board, s) ? '#00c875' : 'var(--wb-subtle)' }} />
                <button type="button" className={styles.linkBtn} style={{ flex: 1 }} onClick={() => props.onOpenItem(s.id)}>{s.title}</button>
              </div>
            ))}
            {rights.canAdd && (
              <input className={styles.addInput} style={{ height: 34, border: '1px dashed var(--wb-border)', borderRadius: 6, marginTop: 6 }}
                placeholder="+ Add subitem" value={newSub} onChange={e => setNewSub(e.target.value)} aria-label="Add subitem"
                onKeyDown={e => {
                  if (e.key === 'Enter' && newSub.trim()) {
                    actions.addItem(item.groupId, newSub.trim(), { parentId: item.id }).catch(() => undefined);
                    setNewSub('');
                  }
                }} />
            )}
          </div>
        )}
      </div>

      <Pivot selectedKey={tab} onLinkClick={i => setTab((i && i.props.itemKey) || 'updates')} styles={{ root: { marginTop: 16 } }}>
        <PivotItem itemKey="updates" headerText="Updates" itemIcon="Comment">
          <UpdatesTab board={board} item={item} canPost={rights.canView} onCountChange={d => actions.bumpUpdateCount(item.id, d)} />
        </PivotItem>
        <PivotItem itemKey="files" headerText="Files" itemIcon="Attach">
          <FilesTab board={board} item={item} canEdit={rights.canEdit} onChanged={async () => {
            try {
              actions.applyItem(await services.items.getItem(board, item.id));
            } catch {
              // The table refreshes on its next poll.
            }
          }} />
        </PivotItem>
        <PivotItem itemKey="activity" headerText="Activity" itemIcon="History">
          <ActivityTab board={board} item={item} />
        </PivotItem>
      </Pivot>
    </Panel>
  );
}

/* ---------- Updates ---------- */

function UpdateBody(props: { body: string; meId: number }): JSX.Element {
  return (
    <div className={styles.updateBody}>
      {parseSegments(props.body).map((s, i) =>
        s.kind === 'text' ? <React.Fragment key={i}>{s.text}</React.Fragment>
          : <span key={i} className={`${styles.mention} ${s.id === props.meId ? styles.mentionMe : ''}`}>@{s.name}</span>
      )}
    </div>
  );
}

function Composer(props: { placeholder: string; onSubmit: (body: string) => Promise<void>; autoFocus?: boolean; onCancel?: () => void }): JSX.Element {
  const { services } = useApp();
  const [text, setText] = React.useState('');
  // People picked from suggestions; the site user id resolves in the background.
  const chosen = React.useRef<{ title: string; person: Promise<IPerson> }[]>([]);
  const [query, setQuery] = React.useState<string | null>(null);
  const [suggestions, setSuggestions] = React.useState<{ key: string; title: string }[]>([]);
  const [active, setActive] = React.useState(0);
  const [busy, setBusy] = React.useState(false);
  const ref = React.useRef<HTMLTextAreaElement>(null);
  // Caret position to restore after inserting a mention, applied before the next keystroke.
  const caretAfterRender = React.useRef<number | null>(null);
  React.useLayoutEffect(() => {
    if (caretAfterRender.current !== null && ref.current) {
      ref.current.setSelectionRange(caretAfterRender.current, caretAfterRender.current);
      caretAfterRender.current = null;
    }
  });

  React.useEffect(() => {
    if (query === null || query.length < 2) {
      setSuggestions([]);
      return undefined;
    }
    const t = setTimeout(() => {
      services.people.search(query, 6).then(r => { setSuggestions(r); setActive(0); }).catch(() => setSuggestions([]));
    }, 250);
    return () => clearTimeout(t);
  }, [query]);

  const onChange = (value: string): void => {
    setText(value);
    const caret = ref.current ? ref.current.selectionStart : value.length;
    const m = /(^|\s)@([^\s@][^@\n]{0,30})$/.exec(value.substring(0, caret));
    setQuery(m ? m[2] : null);
  };

  const pick = (s: { key: string; title: string }): void => {
    // Insert the name straight away so typing can carry on; resolve the user id in the background.
    const el = ref.current;
    const value = el ? el.value : text;
    const caret = el ? el.selectionStart : value.length;
    const before = value.substring(0, caret).replace(/@([^\s@][^@\n]{0,30})$/, '@' + s.title + ' ');
    const next = before + value.substring(caret);
    setText(next);
    if (!chosen.current.some(c => c.title === s.title)) {
      chosen.current.push({ title: s.title, person: services.people.ensure(s.key) });
    }
    setQuery(null);
    setSuggestions([]);
    caretAfterRender.current = before.length;
    if (el) {
      el.focus();
    }
  };

  const submit = async (): Promise<void> => {
    let body = text.trim();
    if (!body) {
      return;
    }
    setBusy(true);
    try {
      // Turn "@Name" into mention tokens for the people picked from suggestions.
      // A person who could not be resolved stays as plain "@Name" text.
      const resolved = await Promise.all(chosen.current.map(c => c.person.catch(() => null)));
      const people = resolved.filter((p): p is IPerson => p !== null);
      people.sort((a, b) => b.title.length - a.title.length).forEach(p => {
        body = body.split('@' + p.title).join(mentionToken(p.title, p.id));
      });
      await props.onSubmit(body);
      setText('');
      chosen.current.splice(0);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.composer}>
      <textarea
        ref={ref}
        className={styles.textarea}
        placeholder={props.placeholder}
        value={text}
        autoFocus={props.autoFocus}
        onChange={e => onChange(e.target.value)}
        aria-label={props.placeholder}
        onKeyDown={e => {
          if (suggestions.length > 0) {
            if (e.key === 'ArrowDown') { e.preventDefault(); setActive(a => (a + 1) % suggestions.length); return; }
            if (e.key === 'ArrowUp') { e.preventDefault(); setActive(a => (a - 1 + suggestions.length) % suggestions.length); return; }
            if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); pick(suggestions[active]); return; }
            if (e.key === 'Escape') { setSuggestions([]); setQuery(null); return; }
          }
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
            submit().catch(() => undefined);
          }
        }}
      />
      {suggestions.length > 0 && (
        <div className={styles.mentionList} role="listbox" aria-label="People to mention">
          {suggestions.map((s, i) => (
            <button key={s.key} type="button" role="option" aria-selected={i === active}
              className={`${styles.mentionOption} ${i === active ? styles.mentionActive : ''}`}
              onMouseDown={e => { e.preventDefault(); pick(s); }}>
              <span className={styles.avatar} style={{ background: personColor(s.title.length) }}>{initials(s.title)}</span>
              {s.title}
            </button>
          ))}
        </div>
      )}
      <div className={styles.row}>
        <span className={`${styles.small} ${styles.muted}`}>Type @ to mention someone. Ctrl+Enter to post.</span>
        <span className={styles.spacer} />
        {props.onCancel && <DefaultButton text="Cancel" onClick={props.onCancel} />}
        <PrimaryButton text={busy ? 'Posting…' : 'Update'} disabled={busy || !text.trim()} onClick={() => { submit().catch(() => undefined); }} />
      </div>
    </div>
  );
}

function UpdatesTab(props: { board: IBoard; item: IWorkItem; canPost: boolean; onCountChange: (delta: number) => void }): JSX.Element {
  const { services, me } = useApp();
  const [updates, setUpdates] = React.useState<IUpdate[] | null>(null);
  const [error, setError] = React.useState('');
  const [replyTo, setReplyTo] = React.useState<number | null>(null);
  const [confirmDialog, confirm] = useConfirm();

  React.useEffect(() => {
    setUpdates(null);
    services.updates.forItem(props.board, props.item.id).then(setUpdates).catch(e => { setError((e as Error).message); setUpdates([]); });
  }, [props.item.id]);

  const post = async (body: string, parentId: number | null): Promise<void> => {
    setError('');
    try {
      const u = await services.updates.add(props.board, props.item.id, body, parentId);
      setUpdates(list => (list || []).concat([u]));
      props.onCountChange(1);
      setReplyTo(null);
    } catch (e) {
      setError('Your update was not posted: ' + (e as Error).message);
      throw e;
    }
  };

  const like = async (u: IUpdate): Promise<void> => {
    try {
      const next = await services.updates.toggleLike(props.board, u, me.id);
      setUpdates(list => (list || []).map(x => (x.id === u.id ? next : x)));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const remove = async (u: IUpdate): Promise<void> => {
    const ok = await confirm({ title: 'Delete this update?', message: 'It goes to the site recycle bin.', confirmText: 'Delete', danger: true });
    if (!ok) {
      return;
    }
    try {
      await services.updates.remove(props.board, u);
      setUpdates(list => (list || []).filter(x => x.id !== u.id && x.parentId !== u.id));
      props.onCountChange(-1);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const renderUpdate = (u: IUpdate, isReply: boolean): JSX.Element => {
    const liked = u.likedByIds.indexOf(me.id) >= 0;
    return (
      <div key={u.id} className={`${styles.update} ${isReply ? styles.reply : ''}`}>
        <div className={styles.row}>
          {u.author && <Avatar person={u.author} />}
          <strong>{u.author ? u.author.title : 'Someone'}</strong>
          <span className={styles.spacer} />
          <span className={`${styles.small} ${styles.muted}`} title={new Date(u.created).toLocaleString()}>{timeAgo(u.created)}</span>
        </div>
        <UpdateBody body={u.body} meId={me.id} />
        <div className={styles.row}>
          <button type="button" className={styles.linkBtn} onClick={() => { like(u).catch(() => undefined); }} aria-pressed={liked}>
            <Icon iconName={liked ? 'LikeSolid' : 'Like'} /> {u.likedByIds.length > 0 ? u.likedByIds.length : ''} Like
          </button>
          {!isReply && props.canPost && <button type="button" className={styles.linkBtn} onClick={() => setReplyTo(u.id)}><Icon iconName="Reply" /> Reply</button>}
          {u.author && u.author.id === me.id && (
            <button type="button" className={styles.linkBtn} onClick={() => { remove(u).catch(() => undefined); }}><Icon iconName="Delete" /> Delete</button>
          )}
        </div>
        {!isReply && (updates || []).filter(r => r.parentId === u.id).map(r => renderUpdate(r, true))}
        {replyTo === u.id && <Composer placeholder="Write a reply…" autoFocus onSubmit={b => post(b, u.id)} onCancel={() => setReplyTo(null)} />}
      </div>
    );
  };

  const top = (updates || []).filter(u => !u.parentId).reverse();
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, paddingTop: 12 }}>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      {props.canPost && <Composer placeholder="Write an update…" onSubmit={b => post(b, null)} />}
      {updates === null && <Spinner size={SpinnerSize.medium} />}
      {updates !== null && top.length === 0 && <p className={styles.muted}>No updates yet. Post one to share progress, ask a question or mention a colleague.</p>}
      {top.map(u => renderUpdate(u, false))}
      {confirmDialog}
    </div>
  );
}

/* ---------- Files ---------- */

function FilesTab(props: { board: IBoard; item: IWorkItem; canEdit: boolean; onChanged: () => Promise<void> }): JSX.Element {
  const { services } = useApp();
  const [files, setFiles] = React.useState<IAttachment[] | null>(null);
  const [busy, setBusy] = React.useState('');
  const [error, setError] = React.useState('');
  const [over, setOver] = React.useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const [confirmDialog, confirm] = useConfirm();

  const load = (): void => {
    services.items.listAttachments(props.board, props.item.id).then(setFiles).catch(e => { setError((e as Error).message); setFiles([]); });
  };
  React.useEffect(load, [props.item.id]);

  const upload = async (list: FileList | null): Promise<void> => {
    if (!list || list.length === 0) {
      return;
    }
    setError('');
    for (let i = 0; i < list.length; i++) {
      const f = list[i];
      setBusy(`Uploading ${f.name}…`);
      try {
        await services.items.addAttachment(props.board, props.item.id, f);
      } catch (e) {
        setError(`${f.name} was not uploaded: ${(e as Error).message}`);
      }
    }
    setBusy('');
    load();
    await props.onChanged();
  };

  const remove = async (f: IAttachment): Promise<void> => {
    const ok = await confirm({ title: `Delete ${f.fileName}?`, message: 'The file goes to the site recycle bin.', confirmText: 'Delete', danger: true });
    if (!ok) {
      return;
    }
    try {
      await services.items.deleteAttachment(props.board, props.item.id, f.fileName);
      load();
      await props.onChanged();
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, paddingTop: 12 }}>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      {props.canEdit && (
        <div
          className={`${styles.dropZone} ${over ? styles.dropZoneOver : ''}`}
          onDragOver={e => { e.preventDefault(); setOver(true); }}
          onDragLeave={() => setOver(false)}
          onDrop={e => { e.preventDefault(); setOver(false); upload(e.dataTransfer.files).catch(() => undefined); }}
        >
          {busy || (
            <>
              Drop files here or <button type="button" className={styles.linkBtn} onClick={() => inputRef.current && inputRef.current.click()}>browse</button>
              <input ref={inputRef} type="file" multiple hidden onChange={e => { upload(e.target.files).catch(() => undefined); e.target.value = ''; }} />
            </>
          )}
        </div>
      )}
      {files === null && <Spinner size={SpinnerSize.medium} />}
      {files !== null && files.length === 0 && <p className={styles.muted}>No files on this item.</p>}
      {(files || []).map(f => (
        <div key={f.fileName} className={styles.fileRow}>
          <Icon iconName="Page" />
          <a href={f.url} target="_blank" rel="noopener noreferrer" data-interception="off">{f.fileName}</a>
          {props.canEdit && (
            <button type="button" className={styles.iconBtn} onClick={() => { remove(f).catch(() => undefined); }} aria-label={`Delete ${f.fileName}`}>
              <Icon iconName="Delete" />
            </button>
          )}
        </div>
      ))}
      {confirmDialog}
    </div>
  );
}

/* ---------- Activity ---------- */

function ActivityTab(props: { board: IBoard; item: IWorkItem }): JSX.Element {
  const { services } = useApp();
  const [entries, setEntries] = React.useState<IActivityEntry[] | null>(null);
  const [error, setError] = React.useState('');
  React.useEffect(() => {
    setEntries(null);
    services.items.activity(props.board, props.item.id).then(setEntries).catch(e => { setError((e as Error).message); setEntries([]); });
  }, [props.item.id, props.item.modified]);

  return (
    <div style={{ paddingTop: 12 }}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {entries === null && <Spinner size={SpinnerSize.medium} />}
      {(entries || []).map(e => (
        <div key={e.version} className={styles.activityRow}>
          <span className={styles.avatar} style={{ background: personColor(e.who.length) }}>{initials(e.who)}</span>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div className={styles.row}>
              <strong>{e.who}</strong>
              <span className={styles.spacer} />
              <span className={`${styles.small} ${styles.muted}`} title={new Date(e.when).toLocaleString()}>{timeAgo(e.when)}</span>
            </div>
            {e.created && <div className={styles.muted}>Created the item</div>}
            {e.changes.map((c, i) => (
              <div key={i} className={styles.change}>
                <span className={styles.muted}>{c.column}:</span>
                {c.from && <span className={styles.strike}>{c.from}</span>}
                {c.from && <Icon iconName="Forward" style={{ fontSize: 10 }} />}
                <span>{c.to || <em className={styles.muted}>empty</em>}</span>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
