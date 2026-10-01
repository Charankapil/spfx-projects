import * as React from 'react';
import { Dialog, DialogFooter, PrimaryButton, DefaultButton, Pivot, PivotItem, TextField, Icon, Spinner, SpinnerSize, MessageBar, MessageBarType, Checkbox } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { ILink } from '../../models/types';
import { ILibrary, IFolderEntry } from '../../services/FilesService';
import { nameFromUrl } from '../../engine/fieldMap';

/**
 * Pick files from a document library on this site, or paste a link to any file
 * (another site, OneDrive, Teams). Returns links; nothing is copied or uploaded.
 */
export function LinkFileDialog(props: { onLink: (links: ILink[]) => Promise<void>; onClose: () => void }): JSX.Element {
  const { services } = useApp();
  const origin = services.files.origin();
  const [tab, setTab] = React.useState('browse');
  const [libraries, setLibraries] = React.useState<ILibrary[] | null>(null);
  const [path, setPath] = React.useState<{ name: string; url: string }[]>([]);
  const [entries, setEntries] = React.useState<IFolderEntry[] | null>(null);
  const [picked, setPicked] = React.useState<IFolderEntry[]>([]);
  const [url, setUrl] = React.useState('');
  const [name, setName] = React.useState('');
  const [error, setError] = React.useState('');
  const [busy, setBusy] = React.useState(false);

  React.useEffect(() => {
    services.files.libraries().then(setLibraries).catch(e => { setError((e as Error).message); setLibraries([]); });
  }, []);

  const open = (crumbs: { name: string; url: string }[]): void => {
    setPath(crumbs);
    setEntries(null);
    if (crumbs.length === 0) {
      return;
    }
    services.files.folder(crumbs[crumbs.length - 1].url).then(setEntries).catch(e => { setError((e as Error).message); setEntries([]); });
  };

  const togglePick = (f: IFolderEntry): void => {
    setPicked(p => (p.some(x => x.url === f.url) ? p.filter(x => x.url !== f.url) : p.concat([f])));
  };

  const validUrl = /^https:\/\/\S+$/i.test(url.trim());
  const links: ILink[] = tab === 'browse'
    ? picked.map(f => ({ name: f.name, url: origin + encodeURI(f.url) }))
    : validUrl ? [{ name: name.trim() || nameFromUrl(url.trim()), url: url.trim() }] : [];

  const submit = async (): Promise<void> => {
    setBusy(true);
    setError('');
    try {
      await props.onLink(links);
      props.onClose();
    } catch (e) {
      setError((e as Error).message);
      setBusy(false);
    }
  };

  return (
    <Dialog hidden={false} onDismiss={busy ? undefined : props.onClose} minWidth={560} maxWidth={680}
      dialogContentProps={{ title: 'Link a file', subText: 'The file stays where it is. The item gets a link to it, and people open it with their own access.' }}>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      <Pivot selectedKey={tab} onLinkClick={i => setTab((i && i.props.itemKey) || 'browse')}>
        <PivotItem itemKey="browse" headerText="Browse this site" itemIcon="FabricFolder">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, padding: '12px 12px 0 0', minHeight: 260, maxHeight: 360, overflowY: 'auto', overflowX: 'hidden' }}>
            <div className={styles.row} style={{ gap: 4 }}>
              <button type="button" className={styles.linkBtn} onClick={() => open([])}>Libraries</button>
              {path.map((c, i) => (
                <React.Fragment key={c.url}>
                  <Icon iconName="ChevronRight" style={{ fontSize: 10 }} />
                  <button type="button" className={styles.linkBtn} onClick={() => open(path.slice(0, i + 1))}>{c.name}</button>
                </React.Fragment>
              ))}
            </div>
            {path.length === 0 && libraries === null && <Spinner size={SpinnerSize.small} />}
            {path.length === 0 && libraries !== null && libraries.length === 0 && <p className={styles.muted}>No document libraries on this site.</p>}
            {path.length === 0 && (libraries || []).map(l => (
              <button key={l.url} type="button" className={styles.navItem} onClick={() => open([{ name: l.title, url: l.url }])}>
                <Icon iconName="DocLibrary" /> <span className={styles.navLabel}>{l.title}</span>
              </button>
            ))}
            {path.length > 0 && entries === null && <Spinner size={SpinnerSize.small} />}
            {path.length > 0 && entries !== null && entries.length === 0 && <p className={styles.muted}>This folder is empty.</p>}
            {path.length > 0 && (entries || []).map(e => (e.isFolder ? (
              <button key={e.url} type="button" className={styles.navItem} onClick={() => open(path.concat([{ name: e.name, url: e.url }]))}>
                <Icon iconName="FabricFolder" /> <span className={styles.navLabel}>{e.name}</span>
              </button>
            ) : (
              <div key={e.url} className={styles.navItem} style={{ cursor: 'default' }}>
                <Checkbox checked={picked.some(x => x.url === e.url)} onChange={() => togglePick(e)} ariaLabel={`Link ${e.name}`} />
                <Icon iconName="Page" />
                <span className={styles.navLabel} onClick={() => togglePick(e)} style={{ cursor: 'pointer' }}>{e.name}</span>
                {e.modified && <span className={`${styles.small} ${styles.muted}`} style={{ flex: 'none' }}>{new Date(e.modified).toLocaleDateString()}</span>}
              </div>
            )))}
          </div>
          {picked.length > 0 && <span className={`${styles.small} ${styles.muted}`}>{picked.length} file{picked.length === 1 ? '' : 's'} selected</span>}
        </PivotItem>
        <PivotItem itemKey="paste" headerText="Paste a link" itemIcon="Link">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, paddingTop: 12, minHeight: 260 }}>
            <TextField label="Link to the file" placeholder="https://contoso.sharepoint.com/sites/team/Shared Documents/plan.xlsx" value={url}
              onChange={(_, v) => setUrl(v || '')} errorMessage={url && !validUrl ? 'Paste a full https:// link' : undefined}
              description="In SharePoint or OneDrive, use Copy link on the file, then paste it here." autoFocus />
            <TextField label="Name (optional)" placeholder={validUrl ? nameFromUrl(url.trim()) : 'Shown on the item'} value={name} onChange={(_, v) => setName(v || '')} />
          </div>
        </PivotItem>
      </Pivot>
      <DialogFooter>
        <PrimaryButton text={busy ? 'Linking…' : links.length > 1 ? `Link ${links.length} files` : 'Link'} disabled={busy || links.length === 0}
          onClick={() => { submit().catch(() => undefined); }} />
        <DefaultButton text="Cancel" disabled={busy} onClick={props.onClose} />
      </DialogFooter>
    </Dialog>
  );
}
