import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  DefaultButton,
  Dropdown,
  Icon,
  IconButton,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SearchBox,
  Spinner,
  SpinnerSize,
  Stack,
  TextField
} from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';
import * as strings from 'StoragePulseWebPartStrings';

import { IStoragePulseProps } from './IStoragePulseProps';
import styles from './StoragePulse.module.scss';
import dash from './dashboard/Dashboard.module.scss';
import { formatCompact, formatDate, formatPercent } from './dashboard/format';
import { ErrorBoundary } from './ErrorBoundary';
import { PulseMark } from './PulseMark';
import { StoragePulse } from './StoragePulse';
import { format, thresholdLabel } from './text';
import { canRunScan, isAllowedPerson } from '../services/access';
import { splitByThreshold, THRESHOLD_OPTIONS } from '../services/activity';
import { ICentralIndex, ISiteSummary, siteKey } from '../services/centralIndex';
import { exportSitesCsv } from '../services/ExportService';
import { formatBytes } from '../services/formatBytes';
import { ResultsStore } from '../services/ResultsStore';
import { safeHref } from '../services/safeData';
import { checkSite, findSites, ISiteEntry, SiteCheck } from '../services/siteDirectory';

type SortKey = 'title' | 'storage' | 'inactive' | 'share' | 'scanned';

const PICKER_ROWS = 100;

function readable(url: string): string {
  try {
    return decodeURI(url);
  } catch {
    return url;
  }
}

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function checkMessage(check: SiteCheck, address: string, host: string): string {
  if (check.ok) {
    return '';
  }
  switch (check.reason) {
    case 'access':
      return format(strings.CentralNoAccess, { site: readable(address) });
    case 'notfound':
      return strings.CentralNotFound;
    case 'host':
      return format(strings.CentralWrongHost, { host });
    case 'invalid':
      return strings.CentralInvalid;
    default:
      return format(strings.CentralCheckFailed, { error: check.detail || '' });
  }
}

interface IRow {
  site: ISiteSummary;
  files: number;
  bytes: number;
  inactiveBytes: number;
  share: number;
  storage: number;
}

/**
 * The central overview: every site collection scanned from this page, and a
 * picker to scan any other one the person can open. Each site opens in the
 * full single-site view, pointed at that site; results are kept here.
 */
export const CentralPulse: React.FC<IStoragePulseProps> = (props) => {
  const storeRef = useRef<ResultsStore>();
  if (!storeRef.current) {
    storeRef.current = new ResultsStore(props.context);
  }
  const host = useMemo(() => new URL(props.context.pageContext.web.absoluteUrl).host, [props.context]);
  const origin = useMemo(() => new URL(props.context.pageContext.web.absoluteUrl).origin, [props.context]);

  const isOwner = useMemo(() => {
    const ctx = props.context.pageContext;
    const legacy = ctx.legacyPageContext as { isSiteAdmin?: boolean } | undefined;
    return !!(legacy && legacy.isSiteAdmin) || ctx.web.permissions.hasPermission(SPPermission.manageWeb);
  }, [props.context]);
  const isListed = useMemo(() => {
    const user = props.context.pageContext.user;
    return isAllowedPerson({ email: user.email, loginName: user.loginName }, props.scanAllowedPeople);
  }, [props.context, props.scanAllowedPeople]);
  const canScan = canRunScan(props.scanPermission, isOwner, isListed);

  const [index, setIndex] = useState<ICentralIndex | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [threshold, setThreshold] = useState<number>(props.thresholdMonths);
  const [open, setOpen] = useState<{ url: string; title: string } | undefined>(undefined);
  const [picking, setPicking] = useState(false);
  const [sites, setSites] = useState<ISiteEntry[] | undefined>(undefined);
  const [sitesError, setSitesError] = useState<string | undefined>(undefined);
  const [filter, setFilter] = useState('');
  const [address, setAddress] = useState('');
  const [checking, setChecking] = useState<string | undefined>(undefined);
  const [checkError, setCheckError] = useState<string | undefined>(undefined);
  const [sortKey, setSortKey] = useState<SortKey>('inactive');
  const [descending, setDescending] = useState(true);
  const mounted = useRef(true);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  useEffect(() => setThreshold(props.thresholdMonths), [props.thresholdMonths]);

  const reload = useCallback(() => {
    (storeRef.current as ResultsStore)
      .loadIndex()
      .then((i) => {
        if (mounted.current) {
          setIndex(i);
          setLoadError(undefined);
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setIndex((before) => before || { schemaVersion: 1, sites: [] });
          setLoadError(format(strings.CentralLoadFailed, { error: errorText(err) }));
        }
      });
  }, []);
  useEffect(reload, [reload]);

  const openPicker = useCallback(() => {
    setPicking(true);
    setCheckError(undefined);
    if (sites === undefined) {
      findSites(props.context)
        .then((found) => mounted.current && setSites(found))
        .catch((err) => {
          if (mounted.current) {
            setSites([]);
            setSitesError(format(strings.CentralSitesFailed, { error: errorText(err) }));
          }
        });
    }
  }, [props.context, sites]);

  /** Checks the person can open the site, then shows it. */
  const choose = useCallback(
    (addressToOpen: string) => {
      setChecking(addressToOpen);
      setCheckError(undefined);
      checkSite(props.context, addressToOpen)
        .then((check) => {
          if (!mounted.current) {
            return;
          }
          setChecking(undefined);
          if (check.ok) {
            setPicking(false);
            setAddress('');
            setOpen({ url: check.url, title: check.title });
          } else {
            setCheckError(checkMessage(check, addressToOpen, host));
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setChecking(undefined);
            setCheckError(format(strings.CentralCheckFailed, { error: errorText(err) }));
          }
        });
    },
    [props.context, host]
  );

  const forget = useCallback(
    (site: ISiteSummary) => {
      if (!window.confirm(format(strings.CentralRemoveConfirm, { site: site.title }))) {
        return;
      }
      (storeRef.current as ResultsStore)
        .forgetSite(site.url)
        .then(reload)
        .catch((err) => setLoadError(errorText(err)));
    },
    [reload]
  );

  const rows: IRow[] = useMemo(
    () =>
      (index ? index.sites : []).map((site) => {
        const split = splitByThreshold(site.histogram, threshold);
        return {
          site,
          files: split.totalFiles,
          bytes: split.totalBytes,
          inactiveBytes: split.inactiveBytes,
          share: split.totalBytes ? split.inactiveBytes / split.totalBytes : 0,
          storage: site.siteStorageBytes || 0
        };
      }),
    [index, threshold]
  );
  const sorted = useMemo(() => {
    const copy = rows.slice();
    copy.sort((a, b) => {
      const cmp =
        sortKey === 'title'
          ? a.site.title.localeCompare(b.site.title)
          : sortKey === 'storage'
          ? a.storage - b.storage
          : sortKey === 'share'
          ? a.share - b.share
          : sortKey === 'scanned'
          ? a.site.scanCompletedAt.localeCompare(b.site.scanCompletedAt)
          : a.inactiveBytes - b.inactiveBytes;
      return descending ? -cmp : cmp;
    });
    return copy;
  }, [rows, sortKey, descending]);
  const totals = useMemo(
    () =>
      rows.reduce(
        (t, r) => ({ storage: t.storage + r.storage, files: t.files + r.files, bytes: t.bytes + r.bytes, inactive: t.inactive + r.inactiveBytes }),
        { storage: 0, files: 0, bytes: 0, inactive: 0 }
      ),
    [rows]
  );

  const visibleSites = useMemo(() => {
    const needle = filter.trim().toLowerCase();
    return (sites || []).filter((s) => !needle || s.title.toLowerCase().indexOf(needle) >= 0 || s.url.toLowerCase().indexOf(needle) >= 0);
  }, [sites, filter]);

  const thresholdOptions: IDropdownOption[] = useMemo(() => THRESHOLD_OPTIONS.map((m) => ({ key: m, text: thresholdLabel(m) })), []);
  const rootClass = `${styles.storagePulse} ${props.theme.isDark ? styles.dark : ''}`;
  const rootStyle = props.theme.accent ? ({ '--sp-accent': props.theme.accent } as React.CSSProperties) : undefined;

  if (open) {
    return (
      <StoragePulse
        {...props}
        key={siteKey(open.url)}
        scope="siteCollection"
        targetSite={open}
        onBack={() => {
          setOpen(undefined);
          reload();
        }}
        onSaved={reload}
      />
    );
  }

  const header = (key: SortKey, label: string, numeric = true): JSX.Element => (
    <th className={numeric ? dash.num : undefined} aria-sort={sortKey === key ? (descending ? 'descending' : 'ascending') : 'none'}>
      <button
        className={dash.sortButton}
        onClick={() => {
          if (sortKey === key) {
            setDescending(!descending);
          } else {
            setSortKey(key);
            setDescending(key !== 'title');
          }
        }}
      >
        {label}
        {sortKey === key && <Icon iconName={descending ? 'SortDown' : 'SortUp'} className={dash.sortIcon} />}
      </button>
    </th>
  );

  const period = thresholdLabel(threshold);

  return (
    <div className={rootClass} style={rootStyle}>
      <header className={styles.hero}>
        <span className={styles.aurora} aria-hidden="true" />
        <div className={styles.heroBrand}>
          <PulseMark size={52} />
          <div className={styles.heroText}>
            <div className={styles.heroEyebrow}>
              {strings.DefaultTitle} · {strings.CentralEyebrow}
            </div>
            <h2 className={styles.heroTitle}>{props.title || strings.DefaultTitle}</h2>
            <div className={styles.heroMeta}>
              {index === undefined
                ? strings.LoadingLastScan
                : index.sites.length === 0
                ? strings.CentralNone
                : format(strings.CentralMeta, { count: index.sites.length })}
            </div>
          </div>
        </div>
        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }} className={styles.heroActions}>
          <Dropdown
            label={strings.InactiveAfterLabel}
            ariaLabel={strings.InactiveAfterAria}
            options={thresholdOptions}
            selectedKey={threshold}
            onChange={(_, option) => option && setThreshold(Number(option.key))}
            className={styles.thresholdPicker}
          />
          <PrimaryButton
            text={canScan ? strings.CentralAddSite : strings.CentralOpenSite}
            iconProps={{ iconName: 'Add' }}
            onClick={openPicker}
            className={styles.scanButton}
          />
          <DefaultButton
            text={strings.ExportCsv}
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={!index || index.sites.length === 0}
            onClick={() => index && exportSitesCsv(index, threshold)}
          />
        </Stack>
      </header>

      {loadError && (
        <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setLoadError(undefined)}>
          {loadError}
        </MessageBar>
      )}

      {picking && (
        <section className={styles.mapPanel} aria-label={strings.CentralPickerTitle}>
          <div className={styles.mapHead}>
            <div>
              <h3 className={styles.mapTitle}>{strings.CentralPickerTitle}</h3>
              <p className={styles.mapIntro}>{strings.CentralPickerIntro}</p>
            </div>
            <DefaultButton text={strings.MapClose} iconProps={{ iconName: 'Cancel' }} onClick={() => setPicking(false)} />
          </div>
          {checkError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setCheckError(undefined)}>
              {checkError}
            </MessageBar>
          )}
          <div className={styles.mapToolbar}>
            <TextField
              label={strings.CentralAddressLabel}
              placeholder={strings.CentralAddressPlaceholder}
              value={address}
              onChange={(_, v) => setAddress(v || '')}
              onKeyDown={(e) => e.key === 'Enter' && address.trim() && choose(address)}
              className={styles.mapSearch}
            />
            <DefaultButton
              text={checking === address && address ? strings.CentralChecking : strings.CentralOpen}
              disabled={!address.trim() || !!checking}
              onClick={() => choose(address)}
              styles={{ root: { alignSelf: 'flex-end' } }}
            />
          </div>
          {sitesError && <MessageBar messageBarType={MessageBarType.warning}>{sitesError}</MessageBar>}
          {sites === undefined ? (
            <Spinner size={SpinnerSize.medium} label={strings.CentralLoadingSites} />
          ) : (
            <>
              <div className={styles.mapToolbar}>
                <SearchBox
                  placeholder={strings.CentralFilter}
                  ariaLabel={strings.CentralFilter}
                  value={filter}
                  onChange={(_, v) => setFilter(v || '')}
                  className={styles.mapSearch}
                />
              </div>
              <div className={styles.mapTree} role="list">
                {visibleSites.slice(0, PICKER_ROWS).map((site) => {
                  const known = (index ? index.sites : []).filter((s) => siteKey(s.url) === siteKey(site.url))[0];
                  return (
                    <div key={site.url} role="listitem" className={`${styles.mapRow} ${styles.centralSiteRow}`}>
                      <span className={styles.mapName}>
                        <Icon iconName="Globe" className={styles.mapIcon} />
                        <strong>{site.title}</strong>
                        <span className={styles.mapMeta}> · {readable(site.url)}</span>
                        {known && (
                          <span className={`${styles.mapBadge} ${styles.mapBadgeDone}`}>
                            {format(strings.CentralScannedOn, { date: formatDate(new Date(known.scanCompletedAt)) })}
                          </span>
                        )}
                      </span>
                      <DefaultButton
                        text={checking === site.url ? strings.CentralChecking : strings.CentralOpen}
                        ariaLabel={format(strings.CentralOpenAria, { site: site.title })}
                        disabled={!!checking}
                        onClick={() => choose(site.url)}
                      />
                    </div>
                  );
                })}
                {visibleSites.length === 0 && <div className={styles.mapMeta}>{strings.CentralNoSites}</div>}
              </div>
              {visibleSites.length > PICKER_ROWS && (
                <div className={styles.mapMeta}>{format(strings.CentralShowing, { shown: PICKER_ROWS, total: visibleSites.length })}</div>
              )}
            </>
          )}
        </section>
      )}

      {index === undefined && (
        <div className={styles.skeleton} aria-hidden="true">
          <span />
          <span />
          <span />
        </div>
      )}

      {index && index.sites.length === 0 && !picking && (
        <div className={styles.emptyState}>
          <PulseMark size={72} />
          <h3 className={styles.emptyTitle}>{strings.CentralEmptyTitle}</h3>
          <p className={styles.emptyText}>{canScan ? strings.CentralEmptyText : strings.EmptyNoPermission}</p>
          {canScan && <PrimaryButton text={strings.CentralAddSite} iconProps={{ iconName: 'Add' }} onClick={openPicker} />}
        </div>
      )}

      {index && index.sites.length > 0 && (
        <div className={dash.dashboard}>
          <section className={dash.card} aria-label={strings.CentralOverviewTitle}>
            <div className={dash.cardHeader}>
              <div>
                <h3 className={dash.cardTitle}>{strings.CentralOverviewTitle}</h3>
                <div className={dash.cardSubtitle}>{format(strings.CentralOverviewSubtitle, { period })}</div>
              </div>
            </div>
            <div className={dash.note} role="note">
              <Icon iconName="BarChartVertical" className={dash.insightIcon} />
              <span>
                {format(strings.CentralTotals, {
                  sites: rows.length,
                  storage: formatBytes(totals.storage),
                  inactive: formatBytes(totals.inactive),
                  share: totals.bytes ? formatPercent(totals.inactive, totals.bytes) : '–',
                  files: formatCompact(totals.files),
                  period
                })}
              </span>
            </div>
            <div className={dash.tableWrap}>
              <table className={dash.table}>
                <thead>
                  <tr>
                    {header('title', strings.CentralColSite, false)}
                    {header('storage', strings.CentralColStorage)}
                    <th className={dash.num}>{strings.ColFiles}</th>
                    {header('inactive', strings.ColInactive)}
                    {header('share', strings.ColInactiveShare)}
                    <th className={dash.num}>{strings.CentralColCoverage}</th>
                    {header('scanned', strings.CentralColScanned)}
                    {canScan && <th aria-label={strings.CentralRemove} />}
                  </tr>
                </thead>
                <tbody>
                  {sorted.map((row) => {
                    const s = row.site;
                    const href = safeHref(s.url, origin);
                    return (
                      <tr key={s.url}>
                        <td>
                          <div className={dash.nameCell}>
                            <button
                              type="button"
                              className={dash.sortButton}
                              onClick={() => setOpen({ url: s.url, title: s.title })}
                              aria-label={format(strings.CentralOpenAria, { site: s.title })}
                            >
                              <strong>{s.title}</strong>
                            </button>
                            {s.failures > 0 && (
                              <span className={dash.warnIcon} title={format(strings.CentralFailures, { count: s.failures })}>
                                <Icon iconName="Warning" />
                              </span>
                            )}
                          </div>
                          <div className={dash.subtle}>
                            {href ? (
                              <a href={href} target="_blank" rel="noopener noreferrer" className={dash.nameLink}>
                                {readable(s.url)}
                              </a>
                            ) : (
                              readable(s.url)
                            )}
                          </div>
                        </td>
                        <td className={dash.num}>{s.siteStorageBytes === undefined ? '–' : formatBytes(s.siteStorageBytes)}</td>
                        <td className={dash.num}>{formatCompact(row.files)}</td>
                        <td className={dash.num}>
                          <strong>{formatBytes(row.inactiveBytes)}</strong>
                        </td>
                        <td className={dash.num}>{row.bytes ? formatPercent(row.inactiveBytes, row.bytes) : '–'}</td>
                        <td className={dash.num}>
                          {format(strings.CentralCoverage, { read: s.librariesRead, total: s.libraries })}
                          {s.librariesEstimated > 0 && (
                            <div className={dash.subtle}>{format(strings.CentralCoverageEstimated, { count: s.librariesEstimated })}</div>
                          )}
                        </td>
                        <td className={dash.num}>
                          {formatDate(new Date(s.scanCompletedAt))}
                          {s.scannedBy && <div className={dash.subtle}>{s.scannedBy}</div>}
                        </td>
                        {canScan && (
                          <td>
                            <IconButton
                              iconProps={{ iconName: 'Delete' }}
                              title={strings.CentralRemove}
                              ariaLabel={`${strings.CentralRemove}: ${s.title}`}
                              onClick={() => forget(s)}
                            />
                          </td>
                        )}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <p className={dash.footnote}>{strings.CentralFootnote}</p>
          </section>
        </div>
      )}
    </div>
  );
};

/** Error boundary wrapper, so a bad saved index cannot break the page. */
export const CentralPulseRoot: React.FC<IStoragePulseProps> = (props) => (
  <ErrorBoundary resetKey={props.context}>
    <CentralPulse {...props} />
  </ErrorBoundary>
);
