import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { extensionOf, fileIcon, formatAge } from '../../common/format';
import { errorText } from '../../common/RequestQueue';
import { fileNameOf, folderOf, kqlPhrase, search, underPath } from '../../common/search';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IMyFilesProps {
  title: string;
  rowCount: string;
}

interface IFile {
  name: string;
  path: string;
  folder: string;
  modified: string;
}

interface ISection {
  files?: IFile[];
  error?: string;
}

interface ISaved {
  recent: ISection;
  checkedOut: ISection;
}

const FileList: React.FC<{ title: string; icon: string; section: ISection | undefined; emptyText: string }> = (props) => (
  <div className={styles.card}>
    <h3 className={styles.cardTitle}>
      <Icon iconName={props.icon} aria-hidden="true" /> {props.title}
      {props.section && props.section.files ? ` (${props.section.files.length})` : ''}
    </h3>
    {!props.section && <Spinner size={SpinnerSize.small} />}
    {props.section && props.section.error && <MessageBar messageBarType={MessageBarType.warning}>{props.section.error}</MessageBar>}
    {props.section && props.section.files && props.section.files.length === 0 && <div className={styles.muted}>{props.emptyText}</div>}
    {props.section && props.section.files && props.section.files.length > 0 && (
      <div className={styles.stack}>
        {props.section.files.map((file) => (
          <div key={file.path} className={styles.fileCell}>
            <Icon iconName={fileIcon(extensionOf(file.name))} aria-hidden="true" />
            <div style={{ minWidth: 0 }}>
              <a className={styles.fileName} href={file.path} target="_blank" rel="noopener noreferrer">
                {file.name}
              </a>
              <div className={styles.path}>
                {file.folder} · {formatAge(file.modified)}
              </div>
            </div>
          </div>
        ))}
      </div>
    )}
  </div>
);

export const MyFiles: React.FC<IToolkitComponentProps<IMyFilesProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const rowCount = parseInt(props.properties.rowCount, 10) || 10;
  const key = cacheKey(kit.scope, `myfiles:${kit.userId}:${rowCount}`);

  const [data, setData] = useState<ISaved | undefined>(undefined);
  const [loading, setLoading] = useState(false);

  const load = useCallback(
    (force: boolean) => {
      const hit = force ? undefined : cacheGet<ISaved>('session', key);
      if (hit) {
        setData(hit);
        return;
      }
      if (!kit.userEmail) {
        setData({ recent: { error: 'Your account has no email address, so your files cannot be looked up.' }, checkedOut: { files: [] } });
        return;
      }
      setLoading(true);
      setData(undefined);
      const me = kqlPhrase(kit.userEmail);
      const where = `IsDocument:1 -FileExtension:aspx ${underPath(kit.sp.webUrl)}`;
      const ask = async (kql: string): Promise<ISection> => {
        try {
          const result = await search(kit.sp, {
            kql,
            select: ['Path', 'LastModifiedTime'],
            rowLimit: rowCount,
            sort: 'LastModifiedTime:descending'
          });
          return {
            files: result.rows.map((row) => ({
              name: fileNameOf(row.Path),
              path: row.Path,
              folder: kit.sp.toServerRelative(folderOf(row.Path)),
              modified: row.LastModifiedTime
            }))
          };
        } catch (err) {
          return { error: `This list could not be loaded: ${errorText(err)}` };
        }
      };
      const run = async (): Promise<ISaved> => {
        // Two requests in total, one after the other.
        const checkedOut = await ask(`${where} CheckoutUserOWSUSER:${me}`);
        const recent = await ask(`${where} EditorOWSUSER:${me}`);
        return { recent, checkedOut };
      };
      run()
        .then((saved) => {
          if (!mounted.current) {
            return;
          }
          if (!saved.recent.error && !saved.checkedOut.error) {
            cacheSet('session', key, saved, 5);
          }
          setData(saved);
          setLoading(false);
        })
        .catch(() => {
          if (mounted.current) {
            setLoading(false);
          }
        });
    },
    [key, kit, mounted, rowCount]
  );

  useEffect(() => {
    load(false);
  }, [load]);

  return (
    <Shell
      icon="Contact"
      title={props.properties.title || 'My Files'}
      subtitle="What you have checked out, and what you last worked on, in this site."
      actions={<DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => load(true)} />}
    >
      {!data && !loading && <Empty icon="Contact" text="Nothing to show yet." />}
      {(data || loading) && (
        <div className={styles.grid} style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))' }}>
          <FileList
            title="Checked out to me"
            icon="Lock"
            section={data ? data.checkedOut : undefined}
            emptyText="You have no files checked out. Nobody is waiting on you."
          />
          <FileList
            title="Recently edited by me"
            icon="Recent"
            section={data ? data.recent : undefined}
            emptyText="You have not edited any files in this site recently."
          />
        </div>
      )}
      <div className={styles.small}>Based on the search index, so a change made a moment ago can take a short while to appear.</div>
    </Shell>
  );
};
