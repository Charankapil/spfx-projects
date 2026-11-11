import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { DefaultButton, Icon, Spinner, SpinnerSize, Toggle } from '@fluentui/react';

import { downloadCsv } from '../../common/csv';
import { CancelToken, errorText, isCancelled } from '../../common/RequestQueue';
import { SpClient } from '../../common/SpClient';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { IPerson, PeoplePicker } from '../../common/ui/PeoplePicker';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { getPhoto } from './photos';
import { initials, IProfile, MAX_PEERS, MAX_REPORTS, nearestManagers, ProfileStore } from './profiles';

export interface IOrgChartProps {
  title: string;
  managerLevels: string;
  showPhotos: boolean;
}

interface IView {
  focus: IProfile;
  managers: IProfile[];
  reports: IProfile[];
  peers?: IProfile[];
  levels: number;
}

/** Colours for initials; each passes 4.5:1 contrast against white text. */
const AVATAR_COLOURS = ['#0f6cbd', '#8a3707', '#0b6a0b', '#5c2e91', '#a4262c', '#005b70', '#6b4a00', '#4f6bed'];

function colourFor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = (hash * 31 + name.charCodeAt(i)) | 0;
  }
  return AVATAR_COLOURS[Math.abs(hash) % AVATAR_COLOURS.length];
}

const Avatar: React.FC<{ sp: SpClient; profile: IProfile; size: number; showPhoto: boolean }> = (props) => {
  const { sp, profile, showPhoto } = props;
  const [photo, setPhoto] = useState<string | undefined>(undefined);

  useEffect(() => {
    let active = true;
    setPhoto(undefined);
    if (showPhoto && profile.email) {
      getPhoto(sp, profile.email)
        .then((url) => {
          if (active) {
            setPhoto(url);
          }
        })
        .catch(() => undefined);
    }
    return () => {
      active = false;
    };
  }, [sp, profile.email, showPhoto]);

  const style: React.CSSProperties = { width: props.size, height: props.size, fontSize: Math.round(props.size * 0.38) };
  return photo ? (
    <img className={styles.avatar} style={style} src={photo} alt="" />
  ) : (
    <span className={styles.avatar} style={{ ...style, background: colourFor(profile.name) }} aria-hidden="true">
      {initials(profile.name)}
    </span>
  );
};

const PersonCard: React.FC<{
  sp: SpClient;
  profile: IProfile;
  showPhoto: boolean;
  onSelect: (profile: IProfile) => void;
}> = (props) => {
  const { profile } = props;
  const count = profile.reports.length;
  return (
    <button
      type="button"
      className={styles.personCard}
      onClick={() => props.onSelect(profile)}
      aria-label={`${profile.name}${profile.jobTitle ? `, ${profile.jobTitle}` : ''}. Show in the chart.`}
    >
      <Avatar sp={props.sp} profile={profile} size={40} showPhoto={props.showPhoto} />
      <span className={styles.personText}>
        <span className={styles.personName}>{profile.name}</span>
        {profile.jobTitle && <span className={styles.small}>{profile.jobTitle}</span>}
        {profile.department && <span className={styles.small}>{profile.department}</span>}
      </span>
      {count > 0 && (
        <span className={styles.pill}>
          {count} report{count === 1 ? '' : 's'}
        </span>
      )}
    </button>
  );
};

export const OrgChart: React.FC<IToolkitComponentProps<IOrgChartProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const storeRef = useRef<ProfileStore | undefined>(undefined);
  if (!storeRef.current) {
    storeRef.current = new ProfileStore(kit.sp);
  }
  const store = storeRef.current;
  const configuredLevels = parseInt(props.properties.managerLevels, 10);
  const defaultLevels = isNaN(configuredLevels) ? 2 : configuredLevels;
  const showPhotos = props.properties.showPhotos !== false;

  const [view, setView] = useState<IView | undefined>(undefined);
  const [history, setHistory] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notFound, setNotFound] = useState(false);
  const [showPeers, setShowPeers] = useState(false);
  const [found, setFound] = useState<IPerson | undefined>(undefined);
  const tokenRef = useRef<CancelToken | undefined>(undefined);

  /** Shows a person: their profile (cached or 1 request) and everyone around them (1 batch per 50). */
  const show = useCallback(
    (account: string | undefined, levels: number, withPeers: boolean, pushHistory?: string) => {
      if (tokenRef.current) {
        tokenRef.current.cancel();
      }
      const token = new CancelToken();
      tokenRef.current = token;
      setLoading(true);
      setError(undefined);
      setNotFound(false);

      const run = async (): Promise<IView | undefined> => {
        const focus = account ? await store.one(account, token) : await store.me(token);
        if (!focus) {
          return undefined;
        }
        const managerAccounts = nearestManagers(focus, levels);
        const reportAccounts = focus.reports.slice(0, MAX_REPORTS);
        const peerAccounts = withPeers ? focus.peers.slice(0, MAX_PEERS) : [];
        const people = await store.many(managerAccounts.concat(reportAccounts, peerAccounts), token);
        const pick = (accounts: string[]): IProfile[] =>
          accounts.map((a) => people[a.toLowerCase()]).filter((p): p is IProfile => !!p);
        const byName = (a: IProfile, b: IProfile): number => a.name.localeCompare(b.name);
        return {
          focus,
          managers: pick(managerAccounts),
          reports: pick(reportAccounts).sort(byName),
          peers: withPeers ? pick(peerAccounts).sort(byName) : undefined,
          levels
        };
      };

      run()
        .then((result) => {
          if (!mounted.current || token.cancelled) {
            return;
          }
          setLoading(false);
          if (!result) {
            setNotFound(true);
            return;
          }
          if (pushHistory) {
            setHistory((current) => current.concat([pushHistory]).slice(-20));
          }
          setView(result);
        })
        .catch((err) => {
          if (!mounted.current || isCancelled(err)) {
            return;
          }
          setLoading(false);
          setError(`The org chart could not be loaded: ${errorText(err)}`);
        });
    },
    [mounted, store]
  );

  useEffect(() => {
    show(undefined, defaultLevels, false);
    return () => {
      if (tokenRef.current) {
        tokenRef.current.cancel();
      }
    };
  }, [defaultLevels, show]);

  const select = useCallback(
    (profile: IProfile) => {
      if (view && profile.account === view.focus.account) {
        return;
      }
      show(profile.account, defaultLevels, showPeers, view ? view.focus.account : undefined);
    },
    [defaultLevels, show, showPeers, view]
  );

  const back = useCallback(() => {
    const previous = history[history.length - 1];
    if (!previous) {
      return;
    }
    setHistory((current) => current.slice(0, -1));
    show(previous, defaultLevels, showPeers);
  }, [defaultLevels, history, show, showPeers]);

  const onFind = useCallback(
    (person: IPerson | undefined) => {
      setFound(undefined);
      if (person) {
        show(person.key, defaultLevels, showPeers, view ? view.focus.account : undefined);
      }
    },
    [defaultLevels, show, showPeers, view]
  );

  const togglePeers = useCallback(
    (on: boolean) => {
      setShowPeers(on);
      if (view && on && !view.peers) {
        show(view.focus.account, view.levels, true);
      }
    },
    [show, view]
  );

  const exportTeam = useCallback(() => {
    if (!view) {
      return;
    }
    const row = (p: IProfile, relation: string): string[] => [relation, p.name, p.jobTitle, p.department, p.office, p.email, p.phone];
    downloadCsv(
      'org-chart',
      ['Relation', 'Name', 'Job title', 'Department', 'Office', 'Email', 'Phone'],
      view.managers
        .map((m) => row(m, 'Manager'))
        .concat([row(view.focus, 'Person')], view.reports.map((r) => row(r, 'Direct report')), (view.peers || []).map((p) => row(p, 'Peer')))
    );
  }, [view]);

  const hiddenManagers = view ? view.focus.managers.length - view.managers.length : 0;
  const focus = view ? view.focus : undefined;
  const isMe = !!focus && !!kit.userEmail && focus.email.toLowerCase() === kit.userEmail.toLowerCase();

  return (
    <Shell
      icon="Org"
      title={props.properties.title || 'Org Chart'}
      subtitle="Who works with whom. Select anyone to move the chart to them."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        <>
          <DefaultButton iconProps={{ iconName: 'Back' }} text="Back" disabled={history.length === 0 || loading} onClick={back} />
          <DefaultButton
            iconProps={{ iconName: 'Contact' }}
            text="Me"
            disabled={loading || isMe}
            onClick={() => show(undefined, defaultLevels, showPeers, focus ? focus.account : undefined)}
          />
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={!view || loading} onClick={exportTeam} />
        </>
      }
    >
      <div className={styles.toolbar}>
        <div className={styles.grow} style={{ maxWidth: 420 }}>
          <PeoplePicker sp={kit.sp} label="Find a person" selected={found} onChange={onFind} disabled={loading} />
        </div>
        <Toggle label="Show peers" inlineLabel checked={showPeers} disabled={loading} onChange={(_, checked) => togglePeers(!!checked)} />
      </div>

      {loading && <Spinner size={SpinnerSize.medium} label="Loading the org chart..." />}

      {notFound && !loading && (
        <Empty
          icon="Org"
          text="No profile was found for this person. People without a Microsoft 365 profile, such as guests, do not appear in the org chart."
        />
      )}

      {view && focus && !loading && (
        <div className={styles.orgChart}>
          {hiddenManagers > 0 && (
            <button type="button" className={styles.linkButton} onClick={() => show(focus.account, view.levels + 3, showPeers)}>
              Show {Math.min(hiddenManagers, 3)} more level{Math.min(hiddenManagers, 3) === 1 ? '' : 's'} above
            </button>
          )}

          {view.managers.map((manager) => (
            <div key={manager.account} className={styles.orgLevel}>
              <PersonCard sp={kit.sp} profile={manager} showPhoto={showPhotos} onSelect={select} />
              <span className={styles.orgLine} aria-hidden="true" />
            </div>
          ))}

          <section className={styles.focusCard} aria-label={`${focus.name}${isMe ? ' (you)' : ''}`}>
            <Avatar sp={kit.sp} profile={focus} size={64} showPhoto={showPhotos} />
            <div className={styles.personText}>
              <span className={styles.personName} style={{ fontSize: 18 }}>
                {focus.name}
                {isMe && <span className={styles.small}> (you)</span>}
              </span>
              {focus.jobTitle && <span>{focus.jobTitle}</span>}
              {(focus.department || focus.office) && (
                <span className={styles.small}>{[focus.department, focus.office].filter(Boolean).join(' · ')}</span>
              )}
              <span className={styles.row} style={{ marginTop: 4 }}>
                {focus.email && (
                  <a className={styles.link} href={`mailto:${focus.email}`}>
                    <Icon iconName="Mail" aria-hidden="true" /> {focus.email}
                  </a>
                )}
                {focus.phone && (
                  <a className={styles.link} href={`tel:${focus.phone.replace(/[^\d+]/g, '')}`}>
                    <Icon iconName="Phone" aria-hidden="true" /> {focus.phone}
                  </a>
                )}
              </span>
            </div>
          </section>

          <div className={styles.orgSection}>
            <h3 className={styles.cardTitle}>{view.reports.length > 0 ? `Direct reports (${view.reports.length})` : 'No direct reports'}</h3>
            {view.reports.length > 0 && (
              <div className={styles.personGrid}>
                {view.reports.map((report) => (
                  <PersonCard key={report.account} sp={kit.sp} profile={report} showPhoto={showPhotos} onSelect={select} />
                ))}
              </div>
            )}
            {focus.reports.length > MAX_REPORTS && (
              <div className={styles.small}>The first {MAX_REPORTS} of {focus.reports.length} direct reports are shown.</div>
            )}
            {Math.min(focus.reports.length, MAX_REPORTS) > view.reports.length && (
              <div className={styles.small}>
                {Math.min(focus.reports.length, MAX_REPORTS) - view.reports.length} more listed as reporting here{' '}
                {Math.min(focus.reports.length, MAX_REPORTS) - view.reports.length === 1 ? 'has' : 'have'} no profile, for example
                because they have left.
              </div>
            )}
          </div>

          {view.peers && (
            <div className={styles.orgSection}>
              <h3 className={styles.cardTitle}>{view.peers.length > 0 ? `Peers (${view.peers.length})` : 'No peers'}</h3>
              {view.peers.length > 0 && (
                <div className={styles.personGrid}>
                  {view.peers.map((peer) => (
                    <PersonCard key={peer.account} sp={kit.sp} profile={peer} showPhoto={showPhotos} onSelect={select} />
                  ))}
                </div>
              )}
              {focus.peers.length > MAX_PEERS && <div className={styles.small}>The first {MAX_PEERS} peers are shown.</div>}
            </div>
          )}
        </div>
      )}

      <div className={styles.small}>
        From your organisation&apos;s Microsoft 365 profiles, as SharePoint sees them. A change of manager can take a day to show here.
      </div>
    </Shell>
  );
};
