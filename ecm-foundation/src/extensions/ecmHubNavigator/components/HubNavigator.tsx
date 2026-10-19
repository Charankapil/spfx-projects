import * as React from 'react';
import * as strings from 'EcmHubNavigatorStrings';

import { IHub, INavigationLocation, INavigationModel, INavLink } from '../models/INavigation';
import styles from './HubNavigator.module.scss';
import { ChevronDown, ChevronUp, CloseIcon, LockIcon, MenuIcon, SearchIcon, ShieldIcon } from './Icons';

export interface IHubNavigatorProps {
  model?: INavigationModel;
  error?: string;
  homeTitle: string;
  homeUrl: string;
  /** Short text in the brand badge; defaults to the home title's initials. */
  brandBadge?: string;
  /** Search results page the box submits to; "{q}" is replaced by the query. */
  searchUrlTemplate: string;
  location: INavigationLocation;
  currentSiteTitle: string;
  classification?: string;
}

function initials(title: string): string {
  return title
    .replace(/&/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w.charAt(0))
    .join('')
    .toUpperCase();
}

function hubStyle(hub: IHub): React.CSSProperties {
  return { ['--hub' as string]: hub.color } as React.CSSProperties;
}

const SiteCard = (props: { site: INavLink; hub: IHub; here: boolean }): JSX.Element => (
  <a className={styles.siteCard} href={props.site.url} aria-current={props.here ? 'page' : undefined}>
    <span className={styles.siteBadge} aria-hidden="true">
      {initials(props.site.title)}
    </span>
    <span className={styles.siteText}>
      <span className={styles.siteTitleRow}>
        <span className={styles.siteTitle}>{props.site.title}</span>
        {props.here && <span className={styles.hereTag}>{strings.YouAreHere}</span>}
        {props.site.restricted && (
          <span className={styles.restrictedTag}>
            <LockIcon />
            {strings.Restricted}
          </span>
        )}
      </span>
      {props.site.description && <span className={styles.siteDesc}>{props.site.description}</span>}
    </span>
  </a>
);

const LinkList = (props: { title: string; links: INavLink[] }): JSX.Element | null =>
  props.links.length ? (
    <div className={styles.column}>
      <div className={styles.columnTitle}>{props.title}</div>
      <ul className={styles.linkList}>
        {props.links.map((link) => (
          <li key={link.id}>
            <a href={link.url}>{link.title}</a>
          </li>
        ))}
      </ul>
    </div>
  ) : null;

export const HubNavigator = (props: IHubNavigatorProps): JSX.Element => {
  const { model, location } = props;
  const [openHubId, setOpenHubId] = React.useState<number | undefined>(undefined);
  const [drawerOpen, setDrawerOpen] = React.useState(false);
  const [drawerHubId, setDrawerHubId] = React.useState<number | undefined>(location.hubId);
  const [query, setQuery] = React.useState('');
  const rootRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        setOpenHubId(undefined);
        setDrawerOpen(false);
      }
    };
    const onClick = (e: MouseEvent): void => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) {
        setOpenHubId(undefined);
      }
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('mousedown', onClick);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('mousedown', onClick);
    };
  }, []);

  const hubs = model?.hubs || [];
  const openHub = hubs.filter((h) => h.id === openHubId)[0];
  const currentHub = hubs.filter((h) => h.id === location.hubId)[0];
  const sharedServices = model?.sharedServices || [];
  const headerLinks = model?.headerLinks || [];

  const onSearch = (e: React.FormEvent): void => {
    e.preventDefault();
    if (query.trim()) {
      window.location.href = props.searchUrlTemplate.replace('{q}', encodeURIComponent(query.trim()));
    }
  };

  const searchBox = (
    <form className={styles.search} role="search" onSubmit={onSearch}>
      <SearchIcon />
      <input
        type="search"
        aria-label={strings.SearchPlaceholder}
        placeholder={strings.SearchPlaceholder}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
    </form>
  );

  return (
    <div className={styles.navigator} ref={rootRef}>
      <div className={styles.bar}>
        <div className={styles.topRow}>
          <a className={styles.brand} href={props.homeUrl}>
            <span className={styles.brandBadge} aria-hidden="true">
              {props.brandBadge || initials(props.homeTitle)}
            </span>
            <span className={styles.brandText}>
              <span className={styles.brandTitle}>{props.homeTitle}</span>
              <span className={styles.brandSub}>{strings.HomeSubtitle}</span>
            </span>
          </a>
          <span className={styles.spacer} />
          <div className={styles.desktopOnly}>{searchBox}</div>
          {headerLinks.map((link, i) => (
            <a key={link.id} href={link.url} className={`${i === 0 ? styles.primaryLink : styles.secondaryLink} ${styles.desktopOnly}`}>
              {link.title}
            </a>
          ))}
          <button
            type="button"
            className={`${styles.iconButton} ${styles.mobileOnly}`}
            aria-label={drawerOpen ? strings.CloseMenu : strings.OpenMenu}
            aria-expanded={drawerOpen}
            onClick={() => setDrawerOpen(!drawerOpen)}
          >
            {drawerOpen ? <CloseIcon /> : <MenuIcon />}
          </button>
        </div>

        {props.error && <div className={styles.status}>{strings.LoadFailed}</div>}
        {!props.error && !model && <div className={styles.status}>{strings.Loading}</div>}
        {!props.error && model && !hubs.length && <div className={styles.status}>{strings.NotConfigured}</div>}

        {hubs.length > 0 && (
          <nav className={`${styles.tabs} ${styles.desktopOnly}`} aria-label={strings.HubsNavLabel}>
            {hubs.map((hub) => {
              const open = hub.id === openHubId;
              const current = hub.id === location.hubId;
              return (
                <button
                  key={hub.id}
                  type="button"
                  style={hubStyle(hub)}
                  className={`${styles.tab} ${open ? styles.tabOpen : ''} ${current && !open ? styles.tabCurrent : ''}`}
                  aria-expanded={open}
                  aria-controls="ecm-hub-mega-menu"
                  onClick={() => setOpenHubId(open ? undefined : hub.id)}
                >
                  <span className={styles.tabDot} aria-hidden="true" />
                  {hub.title}
                  <ChevronDown />
                </button>
              );
            })}
          </nav>
        )}

        {openHub && (
          <div id="ecm-hub-mega-menu" className={styles.mega} style={hubStyle(openHub)}>
            <div className={styles.megaBody}>
              <div className={styles.sitesColumn}>
                <div className={styles.megaHead}>
                  <h2>{openHub.title}</h2>
                  {openHub.url && <a href={openHub.url}>{strings.GoToHub} →</a>}
                </div>
                <div className={styles.columnTitle}>{openHub.isSharedServices ? strings.EnterpriseServices : strings.DepartmentSites}</div>
                <div className={styles.siteGrid}>
                  {openHub.sites.map((site) => (
                    <SiteCard key={site.id} site={site} hub={openHub} here={site.id === location.siteId} />
                  ))}
                </div>
              </div>
              <LinkList title={strings.HubResources} links={openHub.resources} />
              {!openHub.isSharedServices && <LinkList title={strings.SharedServicesColumn} links={sharedServices} />}
              {openHub.description && (
                <div className={styles.governance}>
                  <div className={styles.columnTitle}>{strings.Governance}</div>
                  <p>{openHub.description}</p>
                </div>
              )}
            </div>
            <div className={styles.megaFoot}>
              <span className={styles.trimmed}>
                <ShieldIcon />
                {strings.TrimmedNotice}
              </span>
              <span className={styles.spacer} />
              <button type="button" className={styles.iconButton} aria-label={strings.CloseMenu} onClick={() => setOpenHubId(undefined)}>
                <CloseIcon />
              </button>
            </div>
          </div>
        )}
      </div>

      {drawerOpen && (
        <div className={`${styles.drawer} ${styles.mobileOnly}`}>
          <div className={styles.drawerSearch}>{searchBox}</div>
          <div className={styles.columnTitle}>{strings.HubsNavLabel}</div>
          {hubs.map((hub) => {
            const expanded = hub.id === drawerHubId;
            return (
              <div key={hub.id} className={`${styles.drawerHub} ${expanded ? styles.drawerHubOpen : ''}`} style={hubStyle(hub)}>
                <button type="button" aria-expanded={expanded} onClick={() => setDrawerHubId(expanded ? undefined : hub.id)}>
                  <span className={styles.tabDot} aria-hidden="true" />
                  <span className={styles.spacer}>{hub.title}</span>
                  {expanded ? <ChevronUp /> : <ChevronDown />}
                </button>
                {expanded && (
                  <ul>
                    {hub.sites.map((site) => (
                      <li key={site.id}>
                        <a href={site.url} aria-current={site.id === location.siteId ? 'page' : undefined}>
                          {site.title}
                          {site.id === location.siteId && <span className={styles.hereTag}>{strings.YouAreHere}</span>}
                        </a>
                      </li>
                    ))}
                    {hub.url && (
                      <li>
                        <a className={styles.drawerHubLink} href={hub.url}>
                          {strings.GoToHub} →
                        </a>
                      </li>
                    )}
                  </ul>
                )}
              </div>
            );
          })}
          {headerLinks.map((link, i) => (
            <a key={link.id} href={link.url} className={i === 0 ? styles.drawerPrimary : styles.drawerSecondary}>
              {link.title}
            </a>
          ))}
          <div className={styles.drawerNote}>{strings.TrimmedNotice}</div>
        </div>
      )}

      {currentHub && (
        <div className={styles.breadcrumb}>
          <a href={props.homeUrl}>{props.homeTitle}</a>
          <span aria-hidden="true">›</span>
          {currentHub.url ? <a href={currentHub.url}>{currentHub.title}</a> : <span>{currentHub.title}</span>}
          {location.siteId !== undefined && (
            <>
              <span aria-hidden="true">›</span>
              <span className={styles.crumbCurrent}>{props.currentSiteTitle}</span>
            </>
          )}
          <span className={styles.spacer} />
          {props.classification && <span className={styles.classification}>{props.classification}</span>}
        </div>
      )}
    </div>
  );
};
