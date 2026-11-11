import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { DefaultButton, Icon, Pivot, PivotItem, Spinner, SpinnerSize } from '@fluentui/react';

import { downloadCsv } from '../../common/csv';
import { extensionOf, fileIcon, formatDate, formatNumber } from '../../common/format';
import { errorText, httpStatusOf } from '../../common/RequestQueue';
import { fileNameOf, folderOf, search, underPath } from '../../common/search';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { flagOf } from './principals';

export interface IWhoHasAccessProps {
  title: string;
}

interface IMember {
  Id: number;
  Title: string;
  LoginName?: string;
  PrincipalType: number;
  Email?: string;
}

interface IRoleAssignment {
  Member: IMember;
  RoleDefinitionBindings: Array<{ Name: string }>;
}

interface IAccessRow {
  id: number;
  name: string;
  login: string;
  kind: string;
  isGroup: boolean;
  levels: string[];
  flag?: 'external' | 'everyone';
}

interface ISharedFile {
  name: string;
  path: string;
  folder: string;
  modified: string;
  modifiedBy: string;
  anonymous: boolean;
}

interface IListRow {
  Id: string;
  Title: string;
  ItemCount: number;
  BaseTemplate: number;
  HasUniqueRoleAssignments: boolean;
}

const ROLE_QUERY =
  '$select=Member/Id,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name&$expand=Member,RoleDefinitionBindings';

function kindOf(principalType: number): string {
  switch (principalType) {
    case 1:
      return 'Person';
    case 8:
      return 'SharePoint group';
    case 4:
      return 'Security group';
    case 2:
      return 'Distribution list';
    default:
      return 'Other';
  }
}

function toRows(assignments: IRoleAssignment[]): IAccessRow[] {
  return assignments
    .map((a) => ({
      id: a.Member.Id,
      name: a.Member.Title,
      login: a.Member.LoginName || '',
      kind: kindOf(a.Member.PrincipalType),
      isGroup: a.Member.PrincipalType === 8,
      // "Limited Access" is SharePoint bookkeeping, not a level anyone was given.
      levels: (a.RoleDefinitionBindings || []).map((r) => r.Name).filter((n) => n !== 'Limited Access'),
      flag: flagOf(a.Member.LoginName || '')
    }))
    .filter((row) => row.levels.length > 0)
    .sort((a, b) => a.name.localeCompare(b.name));
}

const Flag: React.FC<{ flag?: 'external' | 'everyone' }> = (props) =>
  props.flag ? (
    <span className={`${styles.pill} ${props.flag === 'external' ? styles.pillWarn : styles.pillBad}`}>
      {props.flag === 'external' ? 'External guest' : 'Everyone in the organisation'}
    </span>
  ) : null;

const AccessTable: React.FC<{
  rows: IAccessRow[];
  members: Record<number, IMember[] | 'loading'>;
  onExpand?: (row: IAccessRow) => void;
}> = (props) => (
  <div className={styles.tableWrap}>
    <table className={styles.table}>
      <thead>
        <tr>
          <th>Who</th>
          <th>Type</th>
          <th>Access</th>
        </tr>
      </thead>
      <tbody>
        {props.rows.map((row) => {
          const loaded = props.members[row.id];
          return (
            <React.Fragment key={row.id}>
              <tr>
                <td>
                  <div className={styles.row}>
                    {row.isGroup && props.onExpand ? (
                      <button type="button" className={styles.linkButton} aria-expanded={!!loaded} onClick={() => props.onExpand && props.onExpand(row)}>
                        {row.name}
                      </button>
                    ) : (
                      <span>{row.name}</span>
                    )}
                    <Flag flag={row.flag} />
                  </div>
                </td>
                <td className={styles.nowrap}>{row.kind}</td>
                <td>{row.levels.join(', ')}</td>
              </tr>
              {loaded && (
                <tr>
                  <td colSpan={3}>
                    {loaded === 'loading' ? (
                      <Spinner size={SpinnerSize.small} label="Loading members..." labelPosition="right" />
                    ) : loaded.length === 0 ? (
                      <span className={styles.muted}>This group has no members.</span>
                    ) : (
                      <div className={styles.row}>
                        {loaded.map((member) => (
                          <span key={member.Id} className={`${styles.pill} ${flagOf(member.LoginName || '') ? styles.pillWarn : ''}`} title={member.Email || ''}>
                            {member.Title}
                          </span>
                        ))}
                      </div>
                    )}
                  </td>
                </tr>
              )}
            </React.Fragment>
          );
        })}
      </tbody>
    </table>
  </div>
);

export const WhoHasAccess: React.FC<IToolkitComponentProps<IWhoHasAccessProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();

  const [site, setSite] = useState<IAccessRow[] | undefined>(undefined);
  const [members, setMembers] = useState<Record<number, IMember[] | 'loading'>>({});
  const [shared, setShared] = useState<ISharedFile[] | undefined>(undefined);
  const [sharedTotal, setSharedTotal] = useState(0);
  const [fileAccess, setFileAccess] = useState<Record<string, IAccessRow[] | 'loading'>>({});
  const [lists, setLists] = useState<IListRow[] | undefined>(undefined);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const explain = useCallback((err: unknown, what: string): string => {
    return httpStatusOf(err) === 403
      ? `${what} can only be seen by people who are allowed to manage permissions on this site.`
      : `${what} could not be loaded: ${errorText(err)}`;
  }, []);

  const loadSite = useCallback(() => {
    setLoading(true);
    setError(undefined);
    kit.sp
      .getJson<{ value: IRoleAssignment[] }>(`${kit.sp.webUrl}/_api/web/RoleAssignments?${ROLE_QUERY}`)
      .then((json) => {
        if (mounted.current) {
          setSite(toRows(json.value || []));
          setLoading(false);
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setSite([]);
          setLoading(false);
          setError(explain(err, 'Site permissions'));
        }
      });
  }, [explain, kit, mounted]);

  useEffect(() => {
    loadSite();
  }, [loadSite]);

  const expandGroup = useCallback(
    (row: IAccessRow) => {
      if (members[row.id]) {
        // Second click collapses; the members stay in memory for a third.
        setMembers((current) => {
          const updated = { ...current };
          delete updated[row.id];
          return updated;
        });
        return;
      }
      setMembers((current) => ({ ...current, [row.id]: 'loading' }));
      kit.sp
        .getJson<{ value: IMember[] }>(
          `${kit.sp.webUrl}/_api/web/SiteGroups/GetById(${row.id})/Users?$select=Id,Title,Email,LoginName,PrincipalType&$top=200`
        )
        .then((json) => {
          if (mounted.current) {
            setMembers((current) => ({ ...current, [row.id]: json.value || [] }));
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setMembers((current) => {
              const updated = { ...current };
              delete updated[row.id];
              return updated;
            });
            setError(explain(err, `The members of "${row.name}"`));
          }
        });
    },
    [explain, kit, members, mounted]
  );

  const loadShared = useCallback(() => {
    setLoading(true);
    setError(undefined);
    search(kit.sp, {
      kql: `IsDocument:1 ViewableByExternalUsers:true ${underPath(kit.sp.webUrl)}`,
      select: ['Path', 'LastModifiedTime', 'ModifiedBy', 'ViewableByAnonymousUsers'],
      rowLimit: 100,
      sort: 'LastModifiedTime:descending'
    })
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        setShared(
          result.rows.map((row) => ({
            name: fileNameOf(row.Path),
            path: row.Path,
            folder: kit.sp.toServerRelative(folderOf(row.Path)),
            modified: row.LastModifiedTime,
            modifiedBy: row.ModifiedBy,
            anonymous: /^(true|1)$/i.test(row.ViewableByAnonymousUsers || '')
          }))
        );
        setSharedTotal(result.total);
        setLoading(false);
      })
      .catch((err) => {
        if (mounted.current) {
          setShared([]);
          setLoading(false);
          setError(explain(err, 'Externally shared files'));
        }
      });
  }, [explain, kit, mounted]);

  const loadLists = useCallback(() => {
    setLoading(true);
    setError(undefined);
    kit.sp
      .getJson<{ value: IListRow[] }>(
        `${kit.sp.webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,ItemCount,BaseTemplate,HasUniqueRoleAssignments`
      )
      .then((json) => {
        if (mounted.current) {
          setLists((json.value || []).filter((l) => l.HasUniqueRoleAssignments).sort((a, b) => a.Title.localeCompare(b.Title)));
          setLoading(false);
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setLists([]);
          setLoading(false);
          setError(explain(err, 'Lists with their own permissions'));
        }
      });
  }, [explain, kit, mounted]);

  const onTab = useCallback(
    (item?: PivotItem) => {
      const key = item ? item.props.itemKey : undefined;
      if (key === 'shared' && !shared && !loading) {
        loadShared();
      }
      if (key === 'lists' && !lists && !loading) {
        loadLists();
      }
    },
    [lists, loadLists, loadShared, loading, shared]
  );

  const showFileAccess = useCallback(
    (file: ISharedFile) => {
      if (fileAccess[file.path]) {
        setFileAccess((current) => {
          const updated = { ...current };
          delete updated[file.path];
          return updated;
        });
        return;
      }
      setFileAccess((current) => ({ ...current, [file.path]: 'loading' }));
      kit.sp
        .getJson<{ value: IRoleAssignment[] }>(
          `${kit.sp.fileApi(kit.sp.toServerRelative(file.path))}/ListItemAllFields/RoleAssignments?${ROLE_QUERY}`
        )
        .then((json) => {
          if (mounted.current) {
            setFileAccess((current) => ({ ...current, [file.path]: toRows(json.value || []) }));
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setFileAccess((current) => {
              const updated = { ...current };
              delete updated[file.path];
              return updated;
            });
            setError(explain(err, `Who can open "${file.name}"`));
          }
        });
    },
    [explain, fileAccess, kit, mounted]
  );

  const exportSite = useCallback(() => {
    downloadCsv(
      'site-access',
      ['Name', 'Type', 'Permission levels', 'Note', 'Login'],
      (site || []).map((r) => [r.name, r.kind, r.levels.join('; '), r.flag || '', r.login])
    );
  }, [site]);

  const flagged = (site || []).filter((r) => r.flag).length;

  return (
    <Shell
      icon="Permissions"
      title={props.properties.title || 'Who Has Access?'}
      subtitle="A plain view of who can reach this site and what is shared outside the organisation."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        <>
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={!site || site.length === 0} onClick={exportSite} />
          <DefaultButton
            iconProps={{ iconName: 'Settings' }}
            text="Manage permissions"
            href={`${kit.sp.webUrl}/_layouts/15/user.aspx`}
            target="_blank"
          />
        </>
      }
    >
      <Pivot onLinkClick={onTab}>
        <PivotItem headerText="This site" itemKey="site">
          <div className={styles.body} style={{ paddingTop: 12 }}>
            {loading && !site && <Spinner size={SpinnerSize.medium} label="Reading site permissions..." />}
            {site && site.length > 0 && (
              <>
                <div className={styles.kpis}>
                  <div className={styles.kpi}>
                    <div className={styles.kpiValue}>{site.length}</div>
                    <div className={styles.kpiLabel}>People and groups with access</div>
                  </div>
                  <div className={styles.kpi}>
                    <div className={styles.kpiValue}>{flagged}</div>
                    <div className={styles.kpiLabel}>Worth a second look (guests, or everyone)</div>
                  </div>
                </div>
                <AccessTable rows={site} members={members} onExpand={expandGroup} />
                <div className={styles.small}>Select a SharePoint group to see its members.</div>
              </>
            )}
            {site && site.length === 0 && !error && <Empty icon="Permissions" text="No permissions were found." />}
          </div>
        </PivotItem>

        <PivotItem headerText="Shared outside" itemKey="shared">
          <div className={styles.body} style={{ paddingTop: 12 }}>
            {loading && !shared && <Spinner size={SpinnerSize.medium} label="Asking the search index..." />}
            {shared && shared.length === 0 && !error && (
              <Empty icon="Completed" text="No files in this site are shared with people outside the organisation." />
            )}
            {shared && shared.length > 0 && (
              <>
                <div className={styles.muted}>
                  {formatNumber(sharedTotal)} file{sharedTotal === 1 ? ' is' : 's are'} visible to people outside the organisation
                  {sharedTotal > shared.length ? `; the ${shared.length} most recently changed are listed` : ''}.
                </div>
                <div className={styles.tableWrap}>
                  <table className={styles.table}>
                    <thead>
                      <tr>
                        <th>File</th>
                        <th>Shared with</th>
                        <th>Modified</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {shared.map((file) => {
                        const access = fileAccess[file.path];
                        return (
                          <React.Fragment key={file.path}>
                            <tr>
                              <td>
                                <div className={styles.fileCell}>
                                  <Icon iconName={fileIcon(extensionOf(file.name))} aria-hidden="true" />
                                  <div>
                                    <a className={styles.fileName} href={file.path} target="_blank" rel="noopener noreferrer">
                                      {file.name}
                                    </a>
                                    <div className={styles.path}>{file.folder}</div>
                                  </div>
                                </div>
                              </td>
                              <td>
                                <span className={`${styles.pill} ${file.anonymous ? styles.pillBad : styles.pillWarn}`}>
                                  {file.anonymous ? 'Anyone with the link' : 'External guests'}
                                </span>
                              </td>
                              <td className={styles.nowrap}>
                                {formatDate(file.modified)}
                                <div className={styles.small}>{file.modifiedBy}</div>
                              </td>
                              <td>
                                <button type="button" className={styles.linkButton} aria-expanded={!!access} onClick={() => showFileAccess(file)}>
                                  {access ? 'Hide access' : 'Who can open it'}
                                </button>
                              </td>
                            </tr>
                            {access && (
                              <tr>
                                <td colSpan={4}>
                                  {access === 'loading' ? (
                                    <Spinner size={SpinnerSize.small} label="Loading..." labelPosition="right" />
                                  ) : access.length === 0 ? (
                                    <span className={styles.muted}>No direct permissions were found on this file.</span>
                                  ) : (
                                    <AccessTable rows={access} members={{}} />
                                  )}
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <div className={styles.small}>
                  To stop sharing a file, open its folder, select the file and choose Manage access. Figures come from the search index.
                </div>
              </>
            )}
          </div>
        </PivotItem>

        <PivotItem headerText="Own permissions" itemKey="lists">
          <div className={styles.body} style={{ paddingTop: 12 }}>
            {loading && !lists && <Spinner size={SpinnerSize.medium} label="Checking lists and libraries..." />}
            {lists && lists.length === 0 && !error && (
              <Empty icon="Completed" text="Every list and library uses the same permissions as the site." />
            )}
            {lists && lists.length > 0 && (
              <>
                <div className={styles.muted}>
                  These lists and libraries no longer follow the site&apos;s permissions, so people may see more or less there than you expect.
                </div>
                <div className={styles.tableWrap}>
                  <table className={styles.table}>
                    <thead>
                      <tr>
                        <th>List or library</th>
                        <th className={styles.num}>Items</th>
                        <th />
                      </tr>
                    </thead>
                    <tbody>
                      {lists.map((list) => (
                        <tr key={list.Id}>
                          <td>{list.Title}</td>
                          <td className={styles.num}>{formatNumber(list.ItemCount)}</td>
                          <td>
                            <a
                              className={styles.link}
                              href={`${kit.sp.webUrl}/_layouts/15/user.aspx?obj=%7B${list.Id}%7D,${list.BaseTemplate === 101 ? 'doclib' : 'list'}&List=%7B${list.Id}%7D`}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              Review permissions
                            </a>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        </PivotItem>
      </Pivot>
    </Shell>
  );
};
