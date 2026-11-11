import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { Checkbox, DefaultButton, Icon, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';

import { cacheGet, cacheKey, cacheRemove, cacheSet } from '../../common/cache';
import { downloadCsv } from '../../common/csv';
import { extensionOf, fileIcon, formatDate } from '../../common/format';
import { field, IListDefinition, isMissingList } from '../../common/lists';
import { errorText } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { SetupCard } from '../../common/ui/SetupCard';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IReadAcknowledgeProps {
  title: string;
  libraryTitle: string;
}

const ACKS: IListDefinition = {
  urlName: 'ToolkitPolicyAcks',
  title: 'Policy Acknowledgements',
  description: 'One row per person per policy version they confirmed reading. Written by the Read & Acknowledge web part.',
  fields: [field.number('PolicyId', 'Policy ID', true), field.text('PolicyVersion', 'Policy Version')],
  indexAuthor: true,
  ownItemsOnly: true
};

interface IPolicy {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Modified: string;
  OData__UIVersionString?: string;
}

interface IAck {
  Id: number;
  PolicyId: number;
  PolicyVersion: string;
  Created: string;
  AuthorId?: number;
  Author?: { Title: string };
}

interface IMember {
  Id: number;
  Title: string;
  PrincipalType: number;
}

/**
 * Only a new major version asks people to confirm again; a typo fix saved as
 * a minor version, or a property change, does not.
 */
function majorVersion(policy: IPolicy): string {
  return String(policy.OData__UIVersionString || '1').split('.')[0];
}

export const ReadAcknowledge: React.FC<IToolkitComponentProps<IReadAcknowledgeProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const libraryTitle = (props.properties.libraryTitle || 'Policies').trim();
  const policiesKey = cacheKey(kit.scope, `policies:${libraryTitle}`);

  const [policies, setPolicies] = useState<IPolicy[] | undefined>(undefined);
  const [myAcks, setMyAcks] = useState<IAck[]>([]);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<string | undefined>(undefined);
  const [ticked, setTicked] = useState<Record<number, boolean>>({});
  const [saving, setSaving] = useState<Record<number, boolean>>({});

  const [report, setReport] = useState<IAck[] | undefined>(undefined);
  const [members, setMembers] = useState<IMember[]>([]);
  const [membersPartial, setMembersPartial] = useState(false);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportCapped, setReportCapped] = useState(false);
  const [openPolicy, setOpenPolicy] = useState<number | undefined>(undefined);

  const load = useCallback(
    (force: boolean) => {
      setLoading(true);
      setError(undefined);
      const run = async (): Promise<{ policies: IPolicy[]; acks: IAck[] }> => {
        let list = force ? undefined : cacheGet<IPolicy[]>('session', policiesKey);
        if (!list) {
          const result = await kit.sp.getAll<IPolicy>(
            `${kit.lists.libraryApi(libraryTitle)}/items?$select=Id,FileLeafRef,FileRef,Modified,OData__UIVersionString` +
              `&$filter=FSObjType eq 0&$top=200`,
            200
          );
          list = result.items;
          cacheSet('session', policiesKey, list, 10);
        }
        const acks = await kit.lists.getItems<IAck>(
          ACKS.urlName,
          `$select=Id,PolicyId,PolicyVersion,Created&$filter=AuthorId eq ${kit.userId}&$top=500`,
          2000
        );
        return { policies: list, acks: acks.items };
      };
      run()
        .then((result) => {
          if (!mounted.current) {
            return;
          }
          result.policies.sort((a, b) => a.FileLeafRef.localeCompare(b.FileLeafRef));
          setPolicies(result.policies);
          setMyAcks(result.acks);
          setNeedsSetup(false);
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setLoading(false);
          if (isMissingList(err)) {
            setNeedsSetup(true);
          } else {
            setError(`The policies could not be loaded: ${errorText(err)}`);
          }
        });
    },
    [kit, libraryTitle, mounted, policiesKey]
  );

  useEffect(() => {
    load(false);
  }, [load]);

  const setup = useCallback(async () => {
    await kit.lists.ensureLibrary(libraryTitle, 'Policies that people are asked to read and acknowledge.');
    await kit.lists.ensure(ACKS);
    cacheRemove('session', policiesKey);
  }, [kit, libraryTitle, policiesKey]);

  const isAcknowledged = useCallback(
    (policy: IPolicy): IAck | undefined =>
      myAcks.filter((a) => a.PolicyId === policy.Id && a.PolicyVersion === majorVersion(policy))[0],
    [myAcks]
  );

  const acknowledge = useCallback(
    (policy: IPolicy) => {
      if (saving[policy.Id]) {
        return;
      }
      setSaving((current) => ({ ...current, [policy.Id]: true }));
      kit.lists
        .addItem<IAck>(ACKS.urlName, {
          Title: policy.FileLeafRef.substring(0, 255),
          PolicyId: policy.Id,
          PolicyVersion: majorVersion(policy)
        })
        .then((created) => {
          if (!mounted.current) {
            return;
          }
          setMyAcks((current) =>
            current.concat([
              { Id: created.Id, PolicyId: policy.Id, PolicyVersion: majorVersion(policy), Created: new Date().toISOString() }
            ])
          );
          setSaving((current) => ({ ...current, [policy.Id]: false }));
          setNotice(`Thank you. Your acknowledgement of "${policy.FileLeafRef}" is recorded.`);
          setReport(undefined);
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setSaving((current) => ({ ...current, [policy.Id]: false }));
          setError(`Your acknowledgement could not be saved: ${errorText(err)}`);
        });
    },
    [kit, mounted, saving]
  );

  const loadReport = useCallback(() => {
    setReportLoading(true);
    setError(undefined);
    const run = async (): Promise<{ acks: IAck[]; capped: boolean; members: IMember[] }> => {
      const acks = await kit.lists.getItems<IAck>(
        ACKS.urlName,
        `$select=Id,PolicyId,PolicyVersion,Created,AuthorId,Author/Title&$expand=Author&$top=1000`,
        5000
      );
      let people: IMember[] = [];
      try {
        const group = await kit.sp.getJson<{ value: IMember[] }>(
          `${kit.sp.webUrl}/_api/web/AssociatedMemberGroup/Users?$select=Id,Title,PrincipalType`
        );
        people = group.value || [];
      } catch {
        // No members group (or no right to read it): the report still shows who has acknowledged.
      }
      return { acks: acks.items, capped: acks.more, members: people };
    };
    run()
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        setReport(result.acks);
        setReportCapped(result.capped);
        setMembers(result.members.filter((m) => m.PrincipalType === 1));
        setMembersPartial(result.members.some((m) => m.PrincipalType !== 1));
        setReportLoading(false);
      })
      .catch((err) => {
        if (mounted.current) {
          setReportLoading(false);
          setError(`The report could not be loaded: ${errorText(err)}`);
        }
      });
  }, [kit, mounted]);

  const acksFor = useCallback(
    (policy: IPolicy): IAck[] => {
      const seen: Record<number, boolean> = {};
      return (report || []).filter((a) => {
        if (a.PolicyId !== policy.Id || a.PolicyVersion !== majorVersion(policy)) {
          return false;
        }
        const author = a.AuthorId || 0;
        if (seen[author]) {
          return false;
        }
        seen[author] = true;
        return true;
      });
    },
    [report]
  );

  const exportReport = useCallback(() => {
    const rows: Array<Array<string | number>> = [];
    for (const policy of policies || []) {
      const acks = acksFor(policy);
      const done: Record<number, boolean> = {};
      for (const ack of acks) {
        done[ack.AuthorId || 0] = true;
        rows.push([policy.FileLeafRef, majorVersion(policy), ack.Author ? ack.Author.Title : '', 'Acknowledged', ack.Created]);
      }
      for (const member of members) {
        if (!done[member.Id]) {
          rows.push([policy.FileLeafRef, majorVersion(policy), member.Title, 'Outstanding', '']);
        }
      }
    }
    downloadCsv('policy-acknowledgements', ['Policy', 'Version', 'Person', 'Status', 'Acknowledged on'], rows);
  }, [acksFor, members, policies]);

  const outstanding = (policies || []).filter((p) => !isAcknowledged(p)).length;

  return (
    <Shell
      icon="ReadingMode"
      title={props.properties.title || 'Read & Acknowledge'}
      subtitle={
        policies && !needsSetup
          ? outstanding === 0
            ? 'You are up to date with every policy.'
            : `${outstanding} polic${outstanding === 1 ? 'y needs' : 'ies need'} your acknowledgement.`
          : undefined
      }
      error={error}
      onDismissError={() => setError(undefined)}
      notice={notice}
      onDismissNotice={() => setNotice(undefined)}
      actions={
        !needsSetup && (
          <>
            <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => load(true)} />
            {kit.canManage && (
              <DefaultButton iconProps={{ iconName: 'ReportDocument' }} text="Who has acknowledged" disabled={reportLoading} onClick={loadReport} />
            )}
          </>
        )
      }
    >
      {loading && <Spinner size={SpinnerSize.medium} label="Loading policies..." />}

      {needsSetup && !loading && (
        <SetupCard
          canManage={kit.canManage}
          what={`a "${libraryTitle}" document library (if it does not exist) and a "Policy Acknowledgements" list`}
          onSetup={setup}
          onDone={() => load(true)}
        />
      )}

      {!loading && !needsSetup && policies && policies.length === 0 && (
        <Empty icon="ReadingMode" text={`There are no documents in the "${libraryTitle}" library yet. Upload a policy to get started.`} />
      )}

      {!needsSetup && policies && policies.length > 0 && (
        <div className={styles.stack}>
          {policies.map((policy) => {
            const ack = isAcknowledged(policy);
            const previous = !ack && myAcks.some((a) => a.PolicyId === policy.Id);
            return (
              <div key={policy.Id} className={styles.card}>
                <div className={styles.spread}>
                  <div className={styles.fileCell}>
                    <Icon iconName={fileIcon(extensionOf(policy.FileLeafRef))} aria-hidden="true" />
                    <div>
                      <a className={styles.fileName} href={`${policy.FileRef}?web=1`} target="_blank" rel="noopener noreferrer">
                        {policy.FileLeafRef}
                      </a>
                      <div className={styles.small}>
                        Version {majorVersion(policy)} · updated {formatDate(policy.Modified)}
                      </div>
                    </div>
                  </div>
                  {ack ? (
                    <span className={`${styles.pill} ${styles.pillGood}`}>Acknowledged {formatDate(ack.Created)}</span>
                  ) : (
                    <span className={`${styles.pill} ${styles.pillWarn}`}>{previous ? 'Updated - please read again' : 'Not acknowledged'}</span>
                  )}
                </div>
                {!ack && (
                  <div className={styles.row} style={{ marginTop: 10 }}>
                    <Checkbox
                      label="I have read and understood this document"
                      checked={!!ticked[policy.Id]}
                      onChange={(_, checked) => setTicked((current) => ({ ...current, [policy.Id]: !!checked }))}
                    />
                    <PrimaryButton
                      text={saving[policy.Id] ? 'Saving...' : 'Acknowledge'}
                      disabled={!ticked[policy.Id] || !!saving[policy.Id]}
                      onClick={() => acknowledge(policy)}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {reportLoading && <Spinner size={SpinnerSize.medium} label="Loading the report..." />}

      {report && policies && (
        <div className={styles.card}>
          <div className={styles.spread}>
            <h3 className={styles.cardTitle}>Who has acknowledged</h3>
            <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" onClick={exportReport} />
          </div>
          {reportCapped && <div className={styles.muted}>Only the first 5,000 acknowledgements are counted here.</div>}
          {membersPartial && (
            <div className={styles.muted}>
              The site members group contains groups as well as people, so the &quot;outstanding&quot; names only cover people added individually.
            </div>
          )}
          <div className={styles.tableWrap} style={{ marginTop: 8 }}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Policy</th>
                  <th className={styles.num}>Acknowledged</th>
                  <th className={styles.num}>Outstanding</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {policies.map((policy) => {
                  const acks = acksFor(policy);
                  const done: Record<number, boolean> = {};
                  for (const ack of acks) {
                    done[ack.AuthorId || 0] = true;
                  }
                  const missing = members.filter((m) => !done[m.Id]);
                  const open = openPolicy === policy.Id;
                  return (
                    <React.Fragment key={policy.Id}>
                      <tr>
                        <td>{policy.FileLeafRef}</td>
                        <td className={styles.num}>{acks.length}</td>
                        <td className={styles.num}>{members.length > 0 ? missing.length : '-'}</td>
                        <td>
                          <button type="button" className={styles.linkButton} aria-expanded={open} onClick={() => setOpenPolicy(open ? undefined : policy.Id)}>
                            {open ? 'Hide names' : 'Show names'}
                          </button>
                        </td>
                      </tr>
                      {open && (
                        <tr>
                          <td colSpan={4}>
                            <div className={styles.small}>Acknowledged</div>
                            <div>{acks.length > 0 ? acks.map((a) => (a.Author ? a.Author.Title : 'Unknown')).join(', ') : 'Nobody yet'}</div>
                            {members.length > 0 && (
                              <>
                                <div className={styles.small} style={{ marginTop: 8 }}>
                                  Outstanding
                                </div>
                                <div>{missing.length > 0 ? missing.map((m) => m.Title).join(', ') : 'Nobody'}</div>
                              </>
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
        </div>
      )}
    </Shell>
  );
};
