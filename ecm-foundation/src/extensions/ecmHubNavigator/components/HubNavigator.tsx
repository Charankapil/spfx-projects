import * as React from 'react';
import * as strings from 'EcmHubNavigatorStrings';

import { IEcmBrand, mix } from '../../../common/theme/ecmTheme';
import { IHub, INavigationLocation, INavigationModel, INavLink } from '../models/INavigation';
import styles from './HubNavigator.module.scss';
import { ArrowRight, ChevronDown, ChevronUp, CloseIcon, HubIcon, LockIcon, MenuIcon, PlusIcon, SearchIcon, ShieldIcon } from './Icons';

export interface IHubNavigatorProps {
  model?: INavigationModel;
  error?: string;
  brand: IEcmBrand;
  homeTitle: string;
  homeUrl: string;
  /** Short text in the badge shown when there is no logo; defaults to the home title's initials. */
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

/** Each hub's colour plus the tints the band and the menu need, as CSS variables. */
function hubStyle(hub: IHub): React.CSSProperties {
  return {
    ['--hub' as string]: hub.color,
    ['--hub-light' as string]: mix(hub.color, '#FFFFFF', 0.45),
    ['--hub-soft' as string]: mix(hub.color, '#FFFFFF', 0.93)
  } as React.CSSProperties;
}

function brandStyle(brand: IEcmBrand): React.CSSProperties {
  return {
    ['--brand' as string]: brand.primary,
    ['--brand-hover' as string]: mix(brand.primary, '#000000', 0.14),
    ['--brand-dark' as string]: brand.dark,
    ['--brand-dark-raised' as string]: mix(brand.dark, '#FFFFFF', 0.08)
  } as React.CSSProperties;
}

const SiteCard = (props: { site: INavLink; here: boolean }): JSX.Element => (
  <a className={`${styles.siteCard} ${props.here ? styles.siteCardHere : ''}`} href={props.site.url} aria-current={props.here ? 'page' : undefined}>
    <span className={styles.siteBadge} aria-hidden="true">
      {initials(props.site.title)}
    </span>
    <span className={styles.siteText}>
      <span className={styles.siteTitleRow}>
        <span className={styles.siteTitle}>{props.site.title}</span>
        {props.here && <span className={styles.hereTag}>{strings.YouAreHere}</span>}
        {props.site.restricted && (
          <span className={styles.restrictedTag} title={strings.Restricted}>
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
    <div className={styles.linkColumn}>
      <div className={styles.columnTitle}>{props.title}</div>
      <ul className={styles.linkList}>
        {props.links.map((link) => (
          <li key={link.id}>
            <a href={link.url}>
              <span>{link.title}</span>
              <ArrowRight />
            </a>
          </li>
        ))}
      </ul>
    </div>
  ) : null;

export const HubNavigator = (props: IHubNavigatorProps): JSX.Element => {
  const { model, location, brand } = props;
  const [openHubId, setOpenHubId] = React.useState<number | undefined>(undefined);
  const [drawerOpen, setDrawerOpen] = React.useState(false);
  const [drawerHubId, setDrawerHubId] = React.useState<number | undefined>(location.hubId);
  const [query, setQuery] = React.useState('');
  const [logoFailed, setLogoFailed] = React.useState(false);
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

  React.useEffect(() => setLogoFailed(false), [brand.logoUrl]);

  const hubs = model?.hubs || [];
  const openHub = hubs.filter((h) => h.id === openHubId)[0];
  const currentHub = hubs.filter((h) => h.id === location.hubId)[0];
  const sharedServices = model?.sharedServices || [];
  const headerLinks = model?.headerLinks || [];
  const showLogo = !!brand.logoUrl && !logoFailed;

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
    <div className={styles.navigator} ref={rootRef} style={brandStyle(brand)}>
      <div className={styles.accentLine} aria-hidden="true" />
      <div className={styles.topRow}>
        <a className={styles.brand} href={props.homeUrl} aria-label={props.homeTitle}>
          {showLogo ? (
            <img className={styles.logo} src={brand.logoUrl} alt="" onError={() => setLogoFailed(true)} />
          ) : (
            <span className={styles.brandBadge} aria-hidden="true">
              {props.brandBadge || initials(props.homeTitle)}
            </span>
          )}
          <span className={styles.brandDivider} aria-hidden="true" />
          <span className={styles.brandText}>
            <span className={styles.brandTitle}>{props.homeTitle}</span>
            <span className={styles.brandSub}>{strings.HomeSubtitle}</span>
          </span>
        </a>
        <span className={styles.spacer} />
        <div className={`${styles.searchWrap} ${styles.desktopOnly}`}>{searchBox}</div>
        {headerLinks.map((link, i) => (
          <a key={link.id} href={link.url} className={`${i === 0 ? styles.primaryLink : styles.secondaryLink} ${styles.desktopOnly}`}>
            {i === 0 && <PlusIcon />}
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

      <div className={styles.band}>
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
                  className={`${styles.tab} ${open ? styles.tabOpen : ''} ${current ? styles.tabCurrent : ''}`}
                  aria-expanded={open}
                  aria-controls="ecm-hub-mega-menu"
                  onClick={() => setOpenHubId(open ? undefined : hub.id)}
                >
                  <span className={styles.tabIcon} aria-hidden="true">
                    <HubIcon title={hub.title} size={14} />
                  </span>
                  {hub.title}
                  <span className={styles.tabChevron}>
                    <ChevronDown />
                  </span>
                </button>
              );
            })}
          </nav>
        )}

        {openHub && (
          <div id="ecm-hub-mega-menu" className={styles.mega} style={hubStyle(openHub)} role="region" aria-label={openHub.title}>
            <div className={styles.megaGrid}>
              <aside className={styles.hubPanel}>
                <span className={styles.hubPanelIcon} aria-hidden="true">
                  <HubIcon title={openHub.title} size={22} />
                </span>
                <h2>{openHub.title}</h2>
                {openHub.description && (
                  <>
                    <div className={styles.columnTitle}>{strings.Governance}</div>
                    <p>{openHub.description}</p>
                  </>
                )}
                <div className={styles.hubPanelMeta}>
                  {openHub.sites.length} {openHub.isSharedServices ? strings.EnterpriseServices.toLowerCase() : strings.DepartmentSites.toLowerCase()}
                </div>
                {openHub.url && (
                  <a className={styles.hubButton} href={openHub.url}>
                    {strings.GoToHub}
                    <ArrowRight />
                  </a>
                )}
              </aside>
              <div className={styles.megaContent}>
                <div className={styles.columnTitle}>{openHub.isSharedServices ? strings.EnterpriseServices : strings.DepartmentSites}</div>
                <div className={styles.siteGrid}>
                  {openHub.sites.map((site) => (
                    <SiteCard key={site.id} site={site} here={site.id === location.siteId} />
                  ))}
                </div>
                <div className={styles.linkColumns}>
                  <LinkList title={strings.HubResources} links={openHub.resources} />
                  {!openHub.isSharedServices && <LinkList title={strings.SharedServicesColumn} links={sharedServices} />}
                </div>
              </div>
            </div>
            <div className={styles.megaFoot}>
              <span className={styles.trimmed}>
                <ShieldIcon />
                {strings.TrimmedNotice}
              </span>
              <span className={styles.spacer} />
              <button type="button" className={styles.closeButton} onClick={() => setOpenHubId(undefined)}>
                <CloseIcon />
                {strings.CloseMenu}
              </button>
            </div>
          </div>
        )}
      </div>

      {drawerOpen && (
        <div className={`${styles.drawer} ${styles.mobileOnly}`}>
          <div className={styles.drawerSearch}>{searchBox}</div>
          <div className={styles.drawerTitle}>{strings.HubsNavLabel}</div>
          {hubs.map((hub) => {
            const expanded = hub.id === drawerHubId;
            return (
              <div key={hub.id} className={`${styles.drawerHub} ${expanded ? styles.drawerHubOpen : ''}`} style={hubStyle(hub)}>
                <button type="button" aria-expanded={expanded} onClick={() => setDrawerHubId(expanded ? undefined : hub.id)}>
                  <span className={styles.drawerIcon} aria-hidden="true">
                    <HubIcon title={hub.title} size={14} />
                  </span>
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
          <span className={styles.crumbSep} aria-hidden="true">
            /
          </span>
          <span className={styles.crumbHub} style={hubStyle(currentHub)}>
            <span className={styles.crumbDot} aria-hidden="true" />
            {currentHub.url ? <a href={currentHub.url}>{currentHub.title}</a> : currentHub.title}
          </span>
          {location.siteId !== undefined && (
            <>
              <span className={styles.crumbSep} aria-hidden="true">
                /
              </span>
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
