import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DefaultButton, Dropdown, PrimaryButton, ProgressIndicator, Spinner, SpinnerSize } from '@fluentui/react';

import { getLibraries, ILibrary } from '../../common/libraries';
import { errorText, httpStatusOf } from '../../common/RequestQueue';
import { IBatchUpdate } from '../../common/batch';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { display, FieldKind, IColumn, parse } from './cells';

export interface IBulkMetadataEditorProps {
  title: string;
}

interface IFieldRow {
  InternalName: string;
  /** The name the REST API uses for the column; differs from InternalName for some columns. */
  EntityPropertyName: string;
  Title: string;
  TypeAsString: string;
  Choices?: string[];
}

interface IItem {
  Id: number;
  FileLeafRef: string;
  FSObjType: number;
  [fieldName: string]: unknown;
}

/** Cell edits not yet saved, keyed "itemId|fieldName". Values are the text shown in the cell. */
type Edits = Record<string, string>;

const SUPPORTED: FieldKind[] = ['Text', 'Number', 'Currency', 'Choice', 'Boolean', 'DateTime'];
const SKIP_FIELDS = ['ContentType', 'FileLeafRef', 'Attachments', '_ExtendedDescription', 'MediaServiceImageTags'];
const PAGE_SIZE = 100;
const BATCH_SIZE = 50;
const MAX_ITEMS_PER_SAVE = 500;
const MAX_UNDO = 50;

function cellKey(id: number, fieldName: string): string {
  return `${id}|${fieldName}`;
}

export const BulkMetadataEditor: React.FC<IToolkitComponentProps<IBulkMetadataEditorProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();

  const [libraries, setLibraries] = useState<ILibrary[]>([]);
  const [library, setLibrary] = useState<ILibrary | undefined>(undefined);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [items, setItems] = useState<IItem[]>([]);
  const [next, setNext] = useState<string | undefined>(undefined);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<string | undefined>(undefined);
  const [edits, setEdits] = useState<Edits>({});
  const [undo, setUndo] = useState<Edits[]>([]);
  const [active, setActive] = useState<{ row: number; column: string } | undefined>(undefined);
  const [progress, setProgress] = useState<{ done: number; total: number } | undefined>(undefined);
  const seq = useRef(0);
  const lastTyped = useRef<string | undefined>(undefined);

  useEffect(() => {
    getLibraries(kit.sp, kit.scope)
      .then((list) => {
        if (mounted.current) {
          setLibraries(list);
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setError(`The libraries could not be listed: ${errorText(err)}`);
        }
      });
  }, [kit, mounted]);

  const files = useMemo(() => items.filter((item) => item.FSObjType === 0), [items]);

  const openLibrary = useCallback(
    (target: ILibrary) => {
      const current = ++seq.current;
      setLibrary(target);
      setLoading(true);
      setError(undefined);
      setNotice(undefined);
      setItems([]);
      setColumns([]);
      setEdits({});
      setUndo([]);
      setNext(undefined);
      setActive(undefined);

      const listApi = `${kit.sp.webUrl}/_api/web/lists(guid'${target.id}')`;
      const run = async (): Promise<{ columns: IColumn[]; items: IItem[]; next: string | undefined }> => {
        const fieldsUrl = `${listApi}/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
        let fields: { value: IFieldRow[] };
        try {
          fields = await kit.sp.getJson<{ value: IFieldRow[] }>(
            `${fieldsUrl}&$select=InternalName,EntityPropertyName,Title,TypeAsString,Choices`
          );
        } catch (err) {
          if (httpStatusOf(err) !== 400) {
            throw err;
          }
          // "Choices" only exists on choice columns; if selecting it is refused, take the full column data.
          fields = await kit.sp.getJson<{ value: IFieldRow[] }>(fieldsUrl);
        }
        const cols: IColumn[] = (fields.value || [])
          .filter((f) => SUPPORTED.indexOf(f.TypeAsString as FieldKind) >= 0 && SKIP_FIELDS.indexOf(f.InternalName) < 0)
          .slice(0, 12)
          .map((f) => ({
            name: f.EntityPropertyName || f.InternalName,
            title: f.Title,
            kind: f.TypeAsString as FieldKind,
            choices: Array.isArray(f.Choices) ? f.Choices : []
          }));
        const select = ['Id', 'FileLeafRef', 'FSObjType'].concat(cols.map((c) => c.name)).join(',');
        const page = await kit.sp.getPage<IItem>(`${listApi}/items?$select=${select}&$orderby=Id&$top=${PAGE_SIZE}`);
        return { columns: cols, items: page.items, next: page.next };
      };
      run()
        .then((result) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          setColumns(result.columns);
          setItems(result.items);
          setNext(result.next);
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          setLoading(false);
          setError(`The library could not be opened: ${errorText(err)}`);
        });
    },
    [kit, mounted]
  );

  const loadMore = useCallback(() => {
    if (!next) {
      return;
    }
    const current = seq.current;
    setLoading(true);
    kit.sp
      .getPage<IItem>(next)
      .then((page) => {
        if (!mounted.current || current !== seq.current) {
          return;
        }
        setItems((existing) => existing.concat(page.items));
        setNext(page.next);
        setLoading(false);
      })
      .catch((err) => {
        if (!mounted.current || current !== seq.current) {
          return;
        }
        setLoading(false);
        setError(`More files could not be loaded: ${errorText(err)}`);
      });
  }, [kit, mounted, next]);

  /**
   * Applies a set of cell changes as one undoable step. Typing in one cell
   * is a single step, however many keystrokes it takes.
   */
  const applyChanges = useCallback(
    (changes: Array<{ item: IItem; column: IColumn; text: string }>) => {
      if (changes.length === 0) {
        return;
      }
      const updated: Edits = { ...edits };
      for (const change of changes) {
        const key = cellKey(change.item.Id, change.column.name);
        if (change.text === display(change.item[change.column.name], change.column.kind)) {
          delete updated[key];
        } else {
          updated[key] = change.text;
        }
      }
      const typingKey = changes.length === 1 ? cellKey(changes[0].item.Id, changes[0].column.name) : undefined;
      if (!typingKey || typingKey !== lastTyped.current) {
        setUndo((stack) => stack.concat([edits]).slice(-MAX_UNDO));
      }
      lastTyped.current = typingKey;
      setEdits(updated);
    },
    [edits]
  );

  const cellText = useCallback(
    (item: IItem, column: IColumn): string => {
      const key = cellKey(item.Id, column.name);
      return Object.prototype.hasOwnProperty.call(edits, key) ? edits[key] : display(item[column.name], column.kind);
    },
    [edits]
  );

  const fillDown = useCallback(() => {
    if (!active) {
      return;
    }
    const column = columns.filter((c) => c.name === active.column)[0];
    const source = files[active.row];
    if (!column || !source) {
      return;
    }
    const text = cellText(source, column);
    applyChanges(files.slice(active.row + 1).map((item) => ({ item, column, text })));
  }, [active, applyChanges, cellText, columns, files]);

  const onPaste = useCallback(
    (event: React.ClipboardEvent, rowIndex: number, columnIndex: number) => {
      const text = event.clipboardData.getData('text/plain');
      // A single value pastes normally; a block copied from Excel spreads across the grid.
      if (text.indexOf('\n') < 0 && text.indexOf('\t') < 0) {
        return;
      }
      event.preventDefault();
      const lines = text.replace(/\r/g, '').split('\n');
      if (lines.length > 1 && lines[lines.length - 1] === '') {
        lines.pop();
      }
      const changes: Array<{ item: IItem; column: IColumn; text: string }> = [];
      lines.forEach((line, r) => {
        const item = files[rowIndex + r];
        if (!item) {
          return;
        }
        line.split('\t').forEach((value, c) => {
          const column = columns[columnIndex + c];
          if (column) {
            changes.push({ item, column, text: value });
          }
        });
      });
      applyChanges(changes);
    },
    [applyChanges, columns, files]
  );

  const undoLast = useCallback(() => {
    if (undo.length === 0) {
      return;
    }
    lastTyped.current = undefined;
    setEdits(undo[undo.length - 1]);
    setUndo(undo.slice(0, -1));
  }, [undo]);

  const invalid = useMemo(() => {
    const bad: Record<string, boolean> = {};
    for (const key of Object.keys(edits)) {
      const column = columns.filter((c) => c.name === key.substring(key.indexOf('|') + 1))[0];
      if (!column || !parse(edits[key], column).ok) {
        bad[key] = true;
      }
    }
    return bad;
  }, [columns, edits]);

  const dirtyCount = Object.keys(edits).length;
  const invalidCount = Object.keys(invalid).length;

  const save = useCallback(() => {
    if (!library || progress) {
      return;
    }
    const byItem: Record<number, Record<string, unknown>> = {};
    for (const key of Object.keys(edits)) {
      const id = Number(key.substring(0, key.indexOf('|')));
      const fieldName = key.substring(key.indexOf('|') + 1);
      const column = columns.filter((c) => c.name === fieldName)[0];
      const parsed = column ? parse(edits[key], column) : { ok: false };
      if (parsed.ok) {
        byItem[id] = byItem[id] || {};
        byItem[id][fieldName] = parsed.value;
      }
    }
    const ids = Object.keys(byItem).map(Number).slice(0, MAX_ITEMS_PER_SAVE);
    if (ids.length === 0) {
      return;
    }
    const listApi = `${kit.sp.webUrl}/_api/web/lists(guid'${library.id}')`;
    setProgress({ done: 0, total: ids.length });
    setError(undefined);
    setNotice(undefined);

    const run = async (): Promise<{ saved: number[]; failures: string[] }> => {
      const saved: number[] = [];
      const failures: string[] = [];
      for (let start = 0; start < ids.length; start += BATCH_SIZE) {
        const chunk = ids.slice(start, start + BATCH_SIZE);
        const updates: IBatchUpdate[] = chunk.map((id) => ({ url: `${listApi}/items(${id})`, body: byItem[id] }));
        const results = await kit.sp.batchMerge(updates);
        results.forEach((result, index) => {
          if (result.ok) {
            saved.push(chunk[index]);
          } else {
            failures.push(`Item ${chunk[index]}: ${result.error || 'not saved'}`);
          }
        });
        if (mounted.current) {
          setProgress({ done: Math.min(start + BATCH_SIZE, ids.length), total: ids.length });
        }
      }
      return { saved, failures };
    };

    run()
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        const savedSet: Record<number, boolean> = {};
        for (const id of result.saved) {
          savedSet[id] = true;
        }
        // Saved values become the stored values; failed and invalid cells stay highlighted.
        setItems((current) => current.map((item) => (savedSet[item.Id] ? { ...item, ...byItem[item.Id] } : item)));
        setEdits((current) => {
          const remaining: Edits = {};
          for (const key of Object.keys(current)) {
            const id = Number(key.substring(0, key.indexOf('|')));
            const fieldName = key.substring(key.indexOf('|') + 1);
            if (!savedSet[id] || !Object.prototype.hasOwnProperty.call(byItem[id], fieldName)) {
              remaining[key] = current[key];
            }
          }
          return remaining;
        });
        setUndo([]);
        lastTyped.current = undefined;
        setProgress(undefined);
        if (result.saved.length > 0) {
          setNotice(`${result.saved.length} file${result.saved.length === 1 ? '' : 's'} updated.`);
        }
        if (result.failures.length > 0) {
          setError(
            `${result.failures.length} file${result.failures.length === 1 ? ' was' : 's were'} not saved and remain highlighted. ` +
              result.failures.slice(0, 3).join(' ')
          );
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setProgress(undefined);
          setError(`Saving stopped: ${errorText(err)} Your unsaved changes are still in the grid.`);
        }
      });
  }, [columns, edits, kit, library, mounted, progress]);

  const dirtyItemCount = useMemo(() => {
    const seen: Record<string, boolean> = {};
    for (const key of Object.keys(edits)) {
      seen[key.substring(0, key.indexOf('|'))] = true;
    }
    return Object.keys(seen).length;
  }, [edits]);

  return (
    <Shell
      icon="Table"
      title={props.properties.title || 'Bulk Metadata Editor'}
      subtitle="Edit the properties of many files at once. Paste from Excel, fill down, then save."
      error={error}
      onDismissError={() => setError(undefined)}
      notice={notice}
      onDismissNotice={() => setNotice(undefined)}
      actions={
        <>
          <DefaultButton iconProps={{ iconName: 'Undo' }} text="Undo" disabled={undo.length === 0 || !!progress} onClick={undoLast} />
          <DefaultButton iconProps={{ iconName: 'Down' }} text="Fill down" disabled={!active || !!progress} onClick={fillDown} />
          <PrimaryButton
            iconProps={{ iconName: 'Save' }}
            text={dirtyCount > 0 ? `Save ${dirtyCount} change${dirtyCount === 1 ? '' : 's'}` : 'Save'}
            disabled={dirtyCount === 0 || dirtyCount === invalidCount || !!progress}
            onClick={save}
          />
        </>
      }
    >
      <div className={styles.toolbar}>
        <Dropdown
          label="Library"
          placeholder="Choose a library"
          options={libraries.map((l) => ({ key: l.id, text: l.title }))}
          selectedKey={library ? library.id : undefined}
          onChange={(_, option) => {
            const target = option && libraries.filter((l) => l.id === option.key)[0];
            if (target) {
              openLibrary(target);
            }
          }}
          disabled={loading || !!progress || dirtyCount > 0}
          styles={{ root: { minWidth: 260 } }}
        />
        {dirtyCount > 0 && <span className={styles.muted}>Save or undo your changes before switching library.</span>}
      </div>

      {invalidCount > 0 && (
        <div className={styles.muted}>
          {invalidCount} highlighted value{invalidCount === 1 ? ' does' : 's do'} not fit the column (for example text in a number
          column, or a choice that does not exist) and will not be saved.
        </div>
      )}
      {dirtyItemCount > MAX_ITEMS_PER_SAVE && (
        <div className={styles.muted}>
          Each save updates up to {MAX_ITEMS_PER_SAVE} files; press Save again for the rest.
        </div>
      )}

      {progress && (
        <ProgressIndicator
          label={`Saving ${progress.done} of ${progress.total} files`}
          percentComplete={progress.total > 0 ? progress.done / progress.total : 0}
        />
      )}

      {!library && <Empty icon="Table" text="Choose a library to start editing." />}

      {library && !loading && columns.length === 0 && items.length > 0 && (
        <Empty icon="Table" text="This library has no editable text, number, choice, yes/no or date columns." />
      )}
      {library && !loading && files.length === 0 && <Empty icon="Page" text="No files were found in this library." />}

      {files.length > 0 && columns.length > 0 && (
        <div className={styles.tableWrap} style={{ maxHeight: 520, overflowY: 'auto' }}>
          <table className={styles.editGrid}>
            <thead>
              <tr>
                <th>File</th>
                {columns.map((column) => (
                  <th key={column.name}>{column.title}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {files.map((item, rowIndex) => (
                <tr key={item.Id}>
                  <td className={styles.fixedCell}>{item.FileLeafRef}</td>
                  {columns.map((column, columnIndex) => {
                    const key = cellKey(item.Id, column.name);
                    const isDirty = Object.prototype.hasOwnProperty.call(edits, key);
                    const label = `${column.title} of ${item.FileLeafRef}`;
                    const common = {
                      className: styles.editCell,
                      'aria-label': label,
                      'aria-invalid': !!invalid[key],
                      disabled: !!progress,
                      onFocus: () => setActive({ row: rowIndex, column: column.name })
                    };
                    const text = cellText(item, column);
                    return (
                      <td key={column.name} className={isDirty ? styles.dirty : undefined} style={invalid[key] ? { outline: '2px solid #a4262c', outlineOffset: -2 } : undefined}>
                        {column.kind === 'Choice' || column.kind === 'Boolean' ? (
                          <select {...common} value={text} onChange={(e) => applyChanges([{ item, column, text: e.target.value }])}>
                            <option value="" />
                            {(column.kind === 'Boolean' ? ['Yes', 'No'] : column.choices).map((choice) => (
                              <option key={choice} value={choice}>
                                {choice}
                              </option>
                            ))}
                            {text !== '' && (column.kind === 'Boolean' ? ['Yes', 'No'] : column.choices).indexOf(text) < 0 && (
                              <option value={text}>{text}</option>
                            )}
                          </select>
                        ) : (
                          <input
                            {...common}
                            type={column.kind === 'DateTime' ? 'date' : 'text'}
                            inputMode={column.kind === 'Number' || column.kind === 'Currency' ? 'decimal' : undefined}
                            value={text}
                            onChange={(e) => applyChanges([{ item, column, text: e.target.value }])}
                            onPaste={(e) => onPaste(e, rowIndex, columnIndex)}
                          />
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {loading && <Spinner size={SpinnerSize.medium} label="Loading..." />}

      {next && !loading && (
        <div className={styles.footer}>
          <DefaultButton text={`Load ${PAGE_SIZE} more`} disabled={!!progress} onClick={loadMore} />
        </div>
      )}

      {library && (
        <div className={styles.small}>
          Text, number, choice, yes/no and date columns can be edited here (up to 12). Dates use your browser&apos;s time zone. Saving
          creates a new version of each changed file&apos;s properties, exactly as editing them one by one would.
        </div>
      )}
    </Shell>
  );
};
