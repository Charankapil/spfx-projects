import * as React from 'react';
import { ChoiceGroup, DefaultButton, Dropdown, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { IDirectoryObject, IGroupRef } from '../../services/GraphApi';
import { compareMembers } from '../../services/m365Review';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

const MAX_COPY = 500;

type Role = 'members' | 'owners';

/**
 * Compares the members (or owners) of two Microsoft 365 groups and copies the missing ones
 * across (Group.ReadWrite.All). Reads two member lists; copies one person per request.
 */
export const CompareGroupsView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Copy / compare members" hint="See who is in one Microsoft 365 group but not another, and copy the missing people across." />
    <GraphGate scopes={['Group.ReadWrite.All']} feature="Copy / compare members">
      <CompareBody />
    </GraphGate>
  </div>
);

const GroupPicker: React.FC<{ label: string; value?: IGroupRef; onPick: (g?: IGroupRef) => void; disabled?: boolean }> = ({ label, value, onPick, disabled }) => {
  const ctx = useAdmin();
  const [text, setText] = React.useState('');
  const [found, setFound] = React.useState<IGroupRef[] | undefined>();
  const [busy, setBusy] = React.useState(false);
  const search = async (): Promise<void> => {
    setBusy(true);
    try {
      setFound(await ctx.graph.searchGroups(text));
    } catch (e) {
      ctx.notify(graphErrorText(e), 'error');
    } finally {
      setBusy(false);
    }
  };
  return (
    <div style={{ flex: '1 1 300px', minWidth: 0 }}>
      <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
        <TextField label={label} placeholder="Start of the group name" value={text} onChange={(_, v) => setText(v || '')} onKeyDown={(e) => e.key === 'Enter' && text.trim() && search()} disabled={disabled} styles={{ root: { flex: 1 } }} />
        <DefaultButton onClick={search} disabled={disabled || busy || !text.trim()}>
          Find
        </DefaultButton>
      </div>
      {found && (
        <Dropdown
          ariaLabel={`${label}: choose a group`}
          placeholder={found.length ? `${found.length} found: choose one` : 'No group starts with that'}
          selectedKey={value ? value.id : null}
          options={found.map((g) => ({ key: g.id, text: `${g.displayName}${g.mail ? ` (${g.mail})` : ''}` }))}
          onChange={(_, o) => o && onPick(found.filter((g) => g.id === o.key)[0])}
          disabled={disabled || !found.length}
          styles={{ root: { marginTop: 6 } }}
        />
      )}
    </div>
  );
};

const CompareBody: React.FC = () => {
  const ctx = useAdmin();
  const [a, setA] = React.useState<IGroupRef | undefined>();
  const [b, setB] = React.useState<IGroupRef | undefined>();
  const [role, setRole] = React.useState<Role>('members');
  const [result, setResult] = React.useState<{ onlyA: IDirectoryObject[]; onlyB: IDirectoryObject[]; both: IDirectoryObject[] } | undefined>();
  const [busy, setBusy] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();

  const compare = async (): Promise<void> => {
    if (!a || !b) {
      return;
    }
    setError(undefined);
    setBusy('Reading both groups…');
    try {
      const [pa, pb] = await Promise.all([ctx.graph.getGroupPeople(a.id, role), ctx.graph.getGroupPeople(b.id, role)]);
      setResult(compareMembers(pa.people, pb.people));
    } catch (e) {
      setError(graphErrorText(e));
    } finally {
      setBusy(undefined);
    }
  };

  const copy = async (from: IGroupRef, to: IGroupRef, people: IDirectoryObject[]): Promise<void> => {
    const list = people.filter((p) => p.type === 'user').slice(0, MAX_COPY);
    if (!list.length) {
      return;
    }
    const ok = await ctx.confirm({
      title: 'Copy people',
      message: `Add ${list.length} ${role === 'owners' ? 'owner' : 'member'}${list.length === 1 ? '' : 's'} of "${from.displayName}" to "${to.displayName}" as ${role}? They get access to its site, mailbox and team.`,
      confirmText: 'Copy'
    });
    if (!ok) {
      return;
    }
    let done = 0;
    const failed: string[] = [];
    for (let i = 0; i < list.length; i++) {
      setBusy(`Copying… ${i + 1} of ${list.length}`);
      try {
        await ctx.graph.addToGroup(to.id, role, list[i].id);
        done++;
      } catch (e) {
        if (/already exist/i.test(graphErrorText(e))) {
          done++;
        } else {
          failed.push(list[i].displayName);
        }
      }
    }
    ctx.log(`Copy group ${role}`, `${from.displayName} → ${to.displayName}: ${done} added`, failed.length === 0, failed.length ? `Failed: ${failed.join(', ')}` : undefined);
    ctx.notify(failed.length ? `${done} copied; failed: ${failed.slice(0, 5).join(', ')}` : `${done} copied to ${to.displayName}.`, failed.length ? 'error' : 'success');
    await compare();
  };

  const List: React.FC<{ title: string; people: IDirectoryObject[]; action?: React.ReactNode }> = ({ title, people, action }) => (
    <Card title={`${title} (${people.length.toLocaleString()})`} right={action}>
      {people.length === 0 ? (
        <Empty text="Nobody." />
      ) : (
        <ul className={styles.memberList}>
          {people.slice(0, 300).map((p) => (
            <li key={p.id} className={styles.memberRow}>
              <span style={{ minWidth: 0 }}>
                <div>
                  {p.displayName} {(p.userType || '').toLowerCase() === 'guest' && <Pill kind="warning">Guest</Pill>} {p.type !== 'user' && <Pill kind="info">{p.type === 'group' ? 'Group' : 'App'}</Pill>}
                </div>
                <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {p.mail || p.userPrincipalName}
                </div>
              </span>
            </li>
          ))}
        </ul>
      )}
      {people.length > 300 && <p className={styles.muted}>Showing 300; export for all.</p>}
    </Card>
  );

  return (
    <>
      <Card>
        <div className={styles.actions} style={{ alignItems: 'flex-start', flexWrap: 'wrap' }}>
          <GroupPicker label="Group A" value={a} onPick={(g) => { setA(g); setResult(undefined); }} disabled={!!busy} />
          <GroupPicker label="Group B" value={b} onPick={(g) => { setB(g); setResult(undefined); }} disabled={!!busy} />
        </div>
        <div className={styles.actions} style={{ marginTop: 10, alignItems: 'flex-end' }}>
          <ChoiceGroup
            label="Compare"
            selectedKey={role}
            onChange={(_, o) => {
              if (o) {
                setRole(o.key as Role);
                setResult(undefined);
              }
            }}
            options={[
              { key: 'members', text: 'Members' },
              { key: 'owners', text: 'Owners' }
            ]}
            styles={{ flexContainer: { display: 'flex', gap: 16 } }}
            disabled={!!busy}
          />
          <PrimaryButton iconProps={{ iconName: 'Compare' }} onClick={compare} disabled={!a || !b || !!busy || (!!a && !!b && a.id === b.id)}>
            {busy || 'Compare'}
          </PrimaryButton>
        </div>
        {a && b && a.id === b.id && <MessageBar messageBarType={MessageBarType.warning}>Choose two different groups.</MessageBar>}
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
      </Card>
      {result && a && b && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="People" label={`Only in ${a.displayName}`} value={result.onlyA.length.toLocaleString()} tone="brand" />
            <StatTile icon="People" label={`Only in ${b.displayName}`} value={result.onlyB.length.toLocaleString()} />
            <StatTile icon="Link" label="In both" value={result.both.length.toLocaleString()} />
          </div>
          <div className={styles.actions} style={{ marginBottom: 8 }}>
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              onClick={() =>
                downloadCsv(
                  `compare-${role}.csv`,
                  ['Where', 'Name', 'E-mail / UPN', 'Type'],
                  ([] as unknown[][])
                    .concat(result.onlyA.map((p) => [`Only in ${a.displayName}`, p.displayName, p.mail || p.userPrincipalName, p.userType || p.type]))
                    .concat(result.onlyB.map((p) => [`Only in ${b.displayName}`, p.displayName, p.mail || p.userPrincipalName, p.userType || p.type]))
                    .concat(result.both.map((p) => ['In both', p.displayName, p.mail || p.userPrincipalName, p.userType || p.type]))
                )
              }
            >
              Export comparison
            </DefaultButton>
          </div>
          <div className={styles.grid}>
            <List
              title={`Only in ${a.displayName}`}
              people={result.onlyA}
              action={
                <PrimaryButton onClick={() => copy(a, b, result.onlyA)} disabled={!!busy || !result.onlyA.some((p) => p.type === 'user')}>
                  Copy to {b.displayName}
                </PrimaryButton>
              }
            />
            <List
              title={`Only in ${b.displayName}`}
              people={result.onlyB}
              action={
                <PrimaryButton onClick={() => copy(b, a, result.onlyB)} disabled={!!busy || !result.onlyB.some((p) => p.type === 'user')}>
                  Copy to {a.displayName}
                </PrimaryButton>
              }
            />
          </div>
          <List title="In both" people={result.both} />
        </>
      )}
    </>
  );
};
