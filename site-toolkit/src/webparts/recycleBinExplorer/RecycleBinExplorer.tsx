import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DefaultButton, Dropdown, IDropdownOption, SearchBox, Spinner, SpinnerSize, Toggle } from '@fluentui/react';

import { downloadCsv } from '../../common/csv';
import { formatBytes, formatDateTime, formatNumber, parseDate } from '../../common/format';
import { errorText, httpStatusOf } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IRecycleBinExplorerProps {
  title: string;
}

interface IBinItem {
  Id: string;
  Title: string;
  LeafName: string;
  DirName: string;
  ItemType: number;
  ItemState: number;
  Size: string | number;
  DeletedDate: string;
  DeletedByName: string;
}

const TYPE_NAMES: Record<number, string> = {
  1: 'File',
  2: 'File version',
  3: 'List item',
  4: 'List',
  5: 'Folder',
  6: 'Folder',
  7: 'Attachment',
  8: 'Item version',
  10: 'Subsite'
};

const TYPE_OPTIONS: IDropdownOption[] = [
  { key: 0, text: 'All types' },
  { key: 1, text: 'Files' },
  { key: 5, text: 'Folders' },
  { key: 3, text: 'List items' },
  { key: 4, text: 'Lists' }
];

const SELECT = '$select=Id,Title,LeafName,DirName,ItemType,ItemState,Size,DeletedDate,DeletedByName';
const STEPS = [500, 1000, 2000, 5000];
const RETENTION_DAYS = 93;

function daysLeft(deleted: string): number {
  const date = parseDate(deleted);
  if (!date) {
    return 0;
  }
  return Math.max(0, RETENTION_DAYS - Math.floor((Date.now() - date.getTime()) / 86400000));
}

export const RecycleBinExplorer: React.FC<IToolkitComponentProps<IRecycleBinExplorerProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();

  const [items, setItems] = useState<IBinItem[] | undefined>(undefined);
  const [step, setStep] = useState(0);
  const [siteWide, setSiteWide] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<string | undefined>(undefined);
  const [text, setText] = useState('');
  const [type, setType] = useState<number>(0);
  const [restoring, setRestoring] = useState<Record<string, boolean>>({});
  const seq = useRef(0);

  const base = siteWide ? `${kit.sp.siteUrl}/_api/site/RecycleBin` : `${kit.sp.webUrl}/_api/web/RecycleBin`;

  const load = useCallback(
    (stepIndex: number) => {
      const current = ++seq.current;
      setLoading(true);
      setError(undefined);
      const top = STEPS[stepIndex];
      const run = async (): Promise<IBinItem[]> => {
        try {
          const sorted = await kit.sp.getJson<{ value: IBinItem[] }>(`${base}?${SELECT}&$orderby=DeletedDate desc&$top=${top}`);
          return sorted.value || [];
        } catch (err) {
          if (httpStatusOf(err) !== 400) {
            throw err;
          }
          // Sorting on the server was refused: take the page as it comes and sort it here.
          const plain = await kit.sp.getJson<{ value: IBinItem[] }>(`${base}?${SELECT}&$top=${top}`);
          return plain.value || [];
        }
      };
      run()
        .then((rows) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          rows.sort((a, b) => (a.DeletedDate < b.DeletedDate ? 1 : a.DeletedDate > b.DeletedDate ? -1 : 0));
          setItems(rows);
          setStep(stepIndex);
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          setLoading(false);
          setItems([]);
          setError(`The recycle bin could not be read: ${errorText(err)}`);
        });
    },
    [base, kit, mounted]
  );

  useEffect(() => {
    load(0);
  }, [load]);

  const restore = useCallback(
    (item: IBinItem) => {
      if (restoring[item.Id]) {
        return;
      }
      setRestoring((current) => ({ ...current, [item.Id]: true }));
      kit.sp
        .post(`${base}('${item.Id}')/restore()`)
        .then(() => {
          if (!mounted.current) {
            return;
          }
          setItems((current) => (current ? current.filter((i) => i.Id !== item.Id) : current));
          setNotice(`"${item.LeafName || item.Title}" was restored to /${item.DirName}.`);
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setRestoring((current) => ({ ...current, [item.Id]: false }));
          setError(`"${item.LeafName || item.Title}" could not be restored: ${errorText(err)}`);
        });
    },
    [base, kit, mounted, restoring]
  );

  const visible = useMemo(() => {
    const needle = text.trim().toLowerCase();
    return (items || []).filter((item) => {
      if (type !== 0 && item.ItemType !== type && !(type === 5 && item.ItemType === 6)) {
        return false;
      }
      if (!needle) {
        return true;
      }
      return (
        `${item.LeafName} ${item.Title} ${item.DirName} ${item.DeletedByName}`.toLowerCase().indexOf(needle) >= 0
      );
    });
  }, [items, text, type]);

  const exportCsv = useCallback(() => {
    downloadCsv(
      'recycle-bin',
      ['Name', 'Original location', 'Type', 'Deleted by', 'Deleted on', 'Size bytes', 'Stage'],
      visible.map((i) => [
        i.LeafName || i.Title,
        `/${i.DirName}`,
        TYPE_NAMES[i.ItemType] || 'Other',
        i.DeletedByName,
        i.DeletedDate,
        Number(i.Size) || 0,
        i.ItemState === 2 ? 'Second stage' : 'First stage'
      ])
    );
  }, [visible]);

  const mayHaveMore = !!items && items.length >= STEPS[step] && step < STEPS.length - 1;

  return (
    <Shell
      icon="RecycleBin"
      title={props.properties.title || 'Recycle Bin Explorer'}
      subtitle="Find out what was deleted, when and by whom, and put it back."
      error={error}
      onDismissError={() => setError(undefined)}
      notice={notice}
      onDismissNotice={() => setNotice(undefined)}
      actions={
        <>
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => load(step)} />
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={visible.length === 0} onClick={exportCsv} />
        </>
      }
    >
      <div className={styles.toolbar}>
        <div className={styles.grow}>
          <SearchBox placeholder="Filter by name, location or person" value={text} onChange={(_, value) => setText(value || '')} />
        </div>
        <Dropdown
          ariaLabel="Type"
          options={TYPE_OPTIONS}
          selectedKey={type}
          onChange={(_, option) => option && setType(option.key as number)}
          styles={{ root: { minWidth: 140 } }}
        />
        {kit.canManage && (
          <Toggle
            label="Whole site collection"
            inlineLabel
            checked={siteWide}
            onChange={(_, checked) => setSiteWide(!!checked)}
            styles={{ root: { marginBottom: 0 } }}
          />
        )}
      </div>

      {loading && <Spinner size={SpinnerSize.medium} label="Reading the recycle bin..." />}

      {!loading && items && visible.length === 0 && (
        <Empty icon="RecycleBin" text={items.length === 0 ? 'The recycle bin is empty.' : 'Nothing in the recycle bin matches that filter.'} />
      )}

      {visible.length > 0 && (
        <>
          <div className={styles.small}>
            Showing {formatNumber(Math.min(visible.length, 500))} of {formatNumber(visible.length)} matching items
            {items && visible.length !== items.length ? ` (${formatNumber(items.length)} loaded)` : ''}.
          </div>
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Type</th>
                  <th>Deleted by</th>
                  <th>Deleted on</th>
                  <th className={styles.num}>Days left</th>
                  <th className={styles.num}>Size</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {visible.slice(0, 500).map((item) => {
                  const left = daysLeft(item.DeletedDate);
                  return (
                    <tr key={item.Id}>
                      <td>
                        <div className={styles.fileName}>{item.LeafName || item.Title}</div>
                        <div className={styles.path}>/{item.DirName}</div>
                      </td>
                      <td className={styles.nowrap}>
                        {TYPE_NAMES[item.ItemType] || 'Other'}
                        {item.ItemState === 2 && <span className={styles.small}> · second stage</span>}
                      </td>
                      <td>{item.DeletedByName}</td>
                      <td className={styles.nowrap}>{formatDateTime(item.DeletedDate)}</td>
                      <td className={styles.num}>
                        <span className={`${styles.pill} ${left <= 7 ? styles.pillBad : left <= 30 ? styles.pillWarn : ''}`}>{left}</span>
                      </td>
                      <td className={styles.num}>{formatBytes(Number(item.Size))}</td>
                      <td>
                        <button type="button" className={styles.linkButton} disabled={!!restoring[item.Id]} onClick={() => restore(item)}>
                          {restoring[item.Id] ? 'Restoring...' : 'Restore'}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {mayHaveMore && !loading && (
        <div className={styles.footer}>
          <DefaultButton text={`Load up to ${formatNumber(STEPS[step + 1])} items`} onClick={() => load(step + 1)} />
        </div>
      )}

      <div className={styles.small}>
        You see the items you are allowed to see: your own deletions, or everything if you manage the site. Deleted items are kept
        for {RETENTION_DAYS} days.
      </div>
    </Shell>
  );
};
