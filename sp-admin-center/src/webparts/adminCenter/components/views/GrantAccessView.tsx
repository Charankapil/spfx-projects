import * as React from 'react';
import { ChoiceGroup, DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, Icon, IChoiceGroupOption, MessageBar, MessageBarType, PrimaryButton, SelectionMode, TextField } from '@fluentui/react';
import { IGroupInfo } from '../../models';
import { AccessResolver, IAssignment, IResolution, IRoleDef, ISecurable, SecurableKind } from '../../services/AccessResolver';
import { parseAddresses } from '../../services/addresses';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { Card, Empty, Loading, Pill, ViewHeader } from '../shared/ui';

const MAX_PEOPLE = 100;

const KIND_TEXT: { [k in SecurableKind]: string } = { site: 'Site collection', web: 'Subsite', library: 'Library', list: 'List', folder: 'Folder', file: 'File', item: 'List item' };
const KIND_ICON: { [k in SecurableKind]: string } = { site: 'SharepointLogo', web: 'Globe', library: 'Library', list: 'BulletedList', folder: 'FabricFolder', file: 'Page', item: 'TextDocument' };

type Who = 'people' | 'addToGroup' | 'group';

interface IObjectData {
  roles: IRoleDef[];
  assignments: IAssignment[];
  siteGroups: IGroupInfo[];
}

/**
 * Paste any address (file, folder, library, list, item, page, site, an "Access denied"
 * page or an Office / sharing link) and grant access on the exact object, or on the
 * parent it inherits its permissions from, to people or to the groups already there.
 */
export const GrantAccessView: React.FC = () => {
  const ctx = useAdmin();
  const resolver = React.useMemo(() => new AccessResolver(ctx.client, ctx.homeWebUrl), [ctx.client, ctx.homeWebUrl]);
  const [input, setInput] = React.useState('');
  const [busy, setBusy] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [res, setRes] = React.useState<IResolution | undefined>();
  const [chosenIndex, setChosenIndex] = React.useState<number | undefined>();
  const [breakHere, setBreakHere] = React.useState(false);
  const [data, setData] = React.useState<IObjectData | undefined>();
  const [dataError, setDataError] = React.useState<string | undefined>();
  const [roleId, setRoleId] = React.useState<number | undefined>();
  const [who, setWho] = React.useState<Who>('people');
  const [people, setPeople] = React.useState('');
  const [groupId, setGroupId] = React.useState<number | undefined>();

  const mounted = React.useRef(true);
  React.useEffect(
    () => () => {
      mounted.current = false;
    },
    []
  );

  const chosen = res && chosenIndex !== undefined ? res.chain[chosenIndex] : undefined;

  // Load permission levels, current access and site groups for the chosen object.
  React.useEffect(() => {
    if (!chosen || !chosen.accessible) {
      setData(undefined);
      return undefined;
    }
    let alive = true;
    setData(undefined);
    setDataError(undefined);
    // Where the chosen object's permissions come from today (itself, or the parent it inherits from).
    const shownFrom = chosen.hasUnique || !res ? chosen : res.chain.slice(0, (chosenIndex as number) + 1).filter((o) => o.hasUnique).pop() || chosen;
    Promise.all([resolver.roleDefinitions(chosen.webUrl), resolver.assignments(shownFrom), ctx.api.getGroups(chosen.webUrl)]).then(
      ([roles, assignments, siteGroups]) => {
        if (!alive) {
          return;
        }
        setData({ roles, assignments, siteGroups });
        setRoleId((cur) => (cur !== undefined && roles.some((r) => r.id === cur) ? cur : (roles.filter((r) => /^(edit|contribute)$/i.test(r.name))[0] || roles[0] || { id: undefined }).id));
      },
      (e: Error) => alive && setDataError(e.message)
    );
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chosen && chosen.api, chosen && chosen.hasUnique, res]);

  const resolve = async (): Promise<void> => {
    setBusy('Finding what this address points to…');
    setError(undefined);
    setRes(undefined);
    setChosenIndex(undefined);
    setBreakHere(false);
    try {
      const r = await resolver.resolve(input);
      if (!mounted.current) {
        return;
      }
      setRes(r);
      // Default: the target itself when it has its own permissions, otherwise where it inherits from.
      const t = r.chain.length - 1;
      const idx = r.chain[t].accessible && r.chain[t].hasUnique ? t : r.source ? r.chain.indexOf(r.source) : t;
      setChosenIndex(idx);
    } catch (e) {
      if (mounted.current) {
        setError((e as Error).message);
      }
    } finally {
      if (mounted.current) {
        setBusy(undefined);
      }
    }
  };

  const role = data && data.roles.filter((r) => r.id === roleId)[0];
  const assignedGroups = data ? data.assignments.filter((a) => a.principalType === 8) : [];
  const groupsWithRole = role ? assignedGroups.filter((a) => a.roles.some((x) => x.toLowerCase() === role.name.toLowerCase())) : [];
  // Adding people to a group changes membership only, so inheritance does not matter for it.
  const willBreak = !!chosen && !chosen.hasUnique && breakHere && who !== 'addToGroup';
  const canAct = !!chosen && chosen.accessible && chosen.canManage !== false && (chosen.hasUnique || breakHere || who === 'addToGroup') && !!role;

  const grant = async (): Promise<void> => {
    if (!chosen || !role || !data || !res) {
      return;
    }
    const addresses = parseAddresses(people).slice(0, MAX_PEOPLE);
    const target = res.target;
    const scopeText = chosen === target ? `${KIND_TEXT[chosen.kind].toLowerCase()} "${chosen.title}"` : `${KIND_TEXT[chosen.kind].toLowerCase()} "${chosen.title}" (and everything that inherits from it, including "${target.title}")`;
    let summary = '';
    if (who === 'people') {
      if (!addresses.length) {
        return;
      }
      summary = `Give ${addresses.length} ${addresses.length === 1 ? 'person' : 'people'} "${role.name}" on the ${scopeText}?`;
    } else if (who === 'addToGroup') {
      const g = assignedGroups.filter((x) => x.principalId === groupId)[0];
      if (!g || !addresses.length) {
        return;
      }
      summary = `Add ${addresses.length} ${addresses.length === 1 ? 'person' : 'people'} to "${g.title}"? They get what that group has everywhere it has access (here: ${g.roles.join(', ') || 'limited access'}).`;
    } else {
      const g = data.siteGroups.filter((x) => x.id === groupId)[0];
      if (!g) {
        return;
      }
      summary = `Give the group "${g.title}" "${role.name}" on the ${scopeText}?`;
    }
    const ok = await ctx.confirm({
      title: 'Grant access',
      message: summary + (willBreak ? ` "${chosen.title}" will first stop inheriting permissions (current permissions are copied).` : ''),
      confirmText: 'Grant'
    });
    if (!ok) {
      return;
    }
    setBusy('Granting…');
    let done = 0;
    const failed: string[] = [];
    try {
      if (who === 'group') {
        await resolver.grant(chosen, groupId as number, role.id, breakHere);
        done = 1;
        ctx.log('Grant access (group)', `${groupId} → ${chosen.path} (${role.name})`, true);
      } else if (who === 'addToGroup') {
        for (const a of addresses) {
          try {
            const u = await ctx.api.ensureUser(chosen.webUrl, a);
            await ctx.api.addUserToGroup(chosen.webUrl, groupId as number, u.loginName);
            ctx.log('Grant access (add to group)', `${u.title} → group ${groupId}`, true);
            done++;
          } catch (e) {
            failed.push(`${a} (${(e as Error).message.substring(0, 100)})`);
            ctx.log('Grant access (add to group)', a, false, (e as Error).message);
          }
        }
      } else {
        for (const a of addresses) {
          try {
            const u = await ctx.api.ensureUser(chosen.webUrl, a);
            await resolver.grant(chosen, u.id, role.id, breakHere);
            ctx.log('Grant access', `${u.title} → ${chosen.path} (${role.name})`, true);
            done++;
          } catch (e) {
            failed.push(`${a} (${(e as Error).message.substring(0, 100)})`);
            ctx.log('Grant access', `${a} → ${chosen.path}`, false, (e as Error).message);
          }
        }
      }
      // Re-read so inheritance flags and current access are up to date before anything else can be clicked.
      setBusy('Refreshing…');
      const r = await resolver.resolve(input).catch(() => undefined);
      if (!mounted.current) {
        return;
      }
      if (r) {
        setRes(r);
      }
      setBreakHere(false);
      setPeople('');
      ctx.notify(failed.length ? `${done} granted. Failed: ${failed.join('; ')}` : `Access granted (${done}). People are not e-mailed automatically; send them the link.`, failed.length ? 'error' : 'success');
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    } finally {
      if (mounted.current) {
        setBusy(undefined);
      }
    }
  };

  const options: IChoiceGroupOption[] = [
    { key: 'people', text: 'Grant specific people this permission level' },
    { key: 'addToGroup', text: 'Add people to a group that has access here' },
    { key: 'group', text: 'Give an existing group this permission level' }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Grant access by URL" hint="Paste any address: a file, folder, library, list, page, site, an Access denied page or an Office / sharing link. The tool finds the exact item and where its permissions come from." />
      <Card>
        <TextField
          label="Address"
          placeholder="https://contoso.sharepoint.com/sites/finance/Shared Documents/Budgets/2026.xlsx"
          value={input}
          onChange={(_, v) => setInput(v || '')}
          onKeyDown={(e) => e.key === 'Enter' && input.trim() && !busy && resolve()}
        />
        <div className={styles.actions} style={{ marginTop: 8 }}>
          <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={resolve} disabled={!input.trim() || !!busy}>
            {busy && !res ? busy : 'Find'}
          </PrimaryButton>
        </div>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {error}
          </MessageBar>
        )}
      </Card>

      {busy && !res && <Loading text={busy} />}

      {res && (
        <>
          {res.notes.length > 0 && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline>
              {res.notes.map((n) => (
                <div key={n}>{n}</div>
              ))}
            </MessageBar>
          )}
          <Card title="Where it is and where its permissions come from">
            <ol className={styles.chain}>
              {res.chain.map((o, i) => (
                <ChainRow
                  key={o.api || o.path + i}
                  o={o}
                  depth={i}
                  isTarget={i === res.chain.length - 1}
                  isSource={o === res.source}
                  chosen={i === chosenIndex}
                  disabled={!!busy}
                  onChoose={() => {
                    setChosenIndex(i);
                    setBreakHere(false);
                  }}
                />
              ))}
            </ol>
            <p className={styles.muted} style={{ marginBottom: 0, fontSize: 12 }}>
              Green edge: the level whose permissions apply to the item you pasted. Choose any level to grant there.
            </p>
          </Card>

          {chosen && (
            <Card title={`Grant on: ${KIND_TEXT[chosen.kind]} "${chosen.title}"`}>
              {!chosen.accessible ? (
                <MessageBar messageBarType={MessageBarType.blocked} isMultiline>
                  You cannot open this {KIND_TEXT[chosen.kind].toLowerCase()}. Ask a site collection administrator, or add yourself as one in SharePoint admin center &gt; Active sites.
                </MessageBar>
              ) : chosen.canManage === false ? (
                <MessageBar messageBarType={MessageBarType.blocked} isMultiline>
                  You do not have the Manage Permissions right here. Choose a level where you do, or ask a site collection administrator.
                </MessageBar>
              ) : (
                <>
                  {!chosen.hasUnique && (
                    <MessageBar messageBarType={MessageBarType.warning} isMultiline>
                      This {KIND_TEXT[chosen.kind].toLowerCase()} <strong>inherits</strong> its permissions{res.source ? ` from "${res.chain.slice(0, (chosenIndex as number) + 1).filter((x) => x.hasUnique).pop()?.title || res.source.title}"` : ''}.
                      Grant on that parent instead (everything under it gets the access), or stop inheriting here.
                      <ChoiceGroup
                        disabled={!!busy}
                        selectedKey={breakHere ? 'break' : 'parent'}
                        onChange={(_, o) => {
                          if (!o) {
                            return;
                          }
                          if (o.key === 'parent') {
                            const parentIdx = res.chain.slice(0, (chosenIndex as number) + 1).map((x, k) => (x.hasUnique ? k : -1)).filter((k) => k >= 0).pop();
                            if (parentIdx !== undefined) {
                              setChosenIndex(parentIdx);
                            }
                            setBreakHere(false);
                          } else {
                            setBreakHere(true);
                          }
                        }}
                        options={[
                          { key: 'parent', text: 'Grant on the parent it inherits from' },
                          { key: 'break', text: `Stop inheriting and grant on this ${KIND_TEXT[chosen.kind].toLowerCase()} only` }
                        ]}
                      />
                    </MessageBar>
                  )}
                  {dataError && <MessageBar messageBarType={MessageBarType.error}>{dataError}</MessageBar>}
                  {!data && !dataError && <Loading text="Loading permission levels and groups…" />}
                  {data && (
                    <div className={styles.grid}>
                      <div>
                        <Dropdown
                          label="Permission level"
                          selectedKey={roleId}
                          onChange={(_, o) => o && setRoleId(Number(o.key))}
                          options={data.roles.map((r) => ({ key: r.id, text: r.name, title: r.description }))}
                        />
                        {role && <p className={styles.muted} style={{ fontSize: 12, margin: '4px 0 8px' }}>{role.description}</p>}
                        <ChoiceGroup label="Grant to" selectedKey={who} onChange={(_, o) => {
                            if (o) {
                              setWho(o.key as Who);
                              setGroupId(undefined);
                            }
                          }} options={options} />
                        {who === 'addToGroup' && (
                          <>
                            <Dropdown
                              label="Group"
                              placeholder="Choose a group"
                              selectedKey={groupId === undefined ? null : groupId}
                              onChange={(_, o) => o && setGroupId(Number(o.key))}
                              options={(groupsWithRole.length ? groupsWithRole : assignedGroups).map((g) => ({ key: g.principalId, text: `${g.title} (${g.roles.join(', ') || 'Limited access'})` }))}
                            />
                            <p className={styles.muted} style={{ fontSize: 12 }}>
                              {role && groupsWithRole.length
                                ? `Groups that have "${role.name}" on this ${KIND_TEXT[chosen.kind].toLowerCase()}.`
                                : role
                                  ? `No group has "${role.name}" here; showing every group that has access.`
                                  : ''}
                            </p>
                          </>
                        )}
                        {who === 'group' && (
                          <Dropdown
                            label="Group"
                            placeholder="Choose a SharePoint group"
                            selectedKey={groupId === undefined ? null : groupId}
                            onChange={(_, o) => o && setGroupId(Number(o.key))}
                            options={data.siteGroups.map((g) => ({ key: g.id, text: g.title }))}
                          />
                        )}
                        {who !== 'group' && (
                          <TextField
                            label="People"
                            placeholder="E-mail addresses, one per line or separated by ; or ,"
                            multiline
                            rows={3}
                            value={people}
                            onChange={(_, v) => setPeople(v || '')}
                          />
                        )}
                        <div className={styles.actions} style={{ marginTop: 10 }}>
                          <PrimaryButton
                            iconProps={{ iconName: 'Permissions' }}
                            onClick={grant}
                            disabled={!canAct || !!busy || (who !== 'people' && groupId === undefined) || (who !== 'group' && !people.trim())}
                          >
                            {busy || 'Grant access'}
                          </PrimaryButton>
                        </div>
                      </div>
                      <div>
                        <h4 style={{ margin: '0 0 6px' }}>Who has access here now{chosen.hasUnique ? '' : ' (inherited)'}</h4>
                        {data.assignments.length === 0 ? (
                          <Empty text="No one is listed." />
                        ) : (
                          <DetailsList
                            items={data.assignments}
                            selectionMode={SelectionMode.none}
                            layoutMode={DetailsListLayoutMode.justified}
                            compact
                            columns={[
                              { key: 'n', name: 'Name', minWidth: 140, isResizable: true, onRender: (a: IAssignment) => a.title },
                              { key: 't', name: 'Type', minWidth: 70, onRender: (a: IAssignment) => (a.principalType === 8 ? 'Group' : a.principalType === 4 ? 'Security group' : 'Person') },
                              { key: 'r', name: 'Permission', minWidth: 110, onRender: (a: IAssignment) => a.roles.join(', ') || 'Limited access' }
                            ]}
                          />
                        )}
                      </div>
                    </div>
                  )}
                </>
              )}
            </Card>
          )}
        </>
      )}
    </div>
  );
};

const ChainRow: React.FC<{ o: ISecurable; depth: number; isTarget: boolean; isSource: boolean; chosen: boolean; disabled: boolean; onChoose: () => void }> = ({ o, depth, isTarget, isSource, chosen, disabled, onChoose }) => (
  <li className={`${styles.chainRow} ${isTarget ? styles.chainRowTarget : ''} ${isSource ? styles.chainRowSource : ''}`} style={{ marginLeft: Math.min(depth, 8) * 14 }}>
    <Icon iconName={KIND_ICON[o.kind]} aria-hidden="true" />
    <span className={styles.chainMain}>
      <div>
        <strong>{o.title}</strong> <span className={styles.muted}>· {KIND_TEXT[o.kind]}</span> {isTarget && <Pill kind="info">Pasted item</Pill>}
      </div>
      <div className={styles.chainPath} title={o.path}>
        {o.path}
      </div>
    </span>
    {!o.accessible ? (
      <Pill kind="critical">No access</Pill>
    ) : (
      <>
        {o.hasUnique ? <Pill kind="good">Unique permissions</Pill> : <Pill>Inherits</Pill>}
        {o.canManage === false && <Pill kind="warning">Can&rsquo;t manage</Pill>}
      </>
    )}
    <DefaultButton onClick={onChoose} disabled={!o.accessible || disabled} checked={chosen} styles={{ root: { minWidth: 90 } }} aria-pressed={chosen}>
      {chosen ? 'Selected' : 'Grant here'}
    </DefaultButton>
  </li>
);
