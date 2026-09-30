import * as React from 'react';
import { useMemo, useState } from 'react';
import { Checkbox, DefaultButton, Icon, MessageBar, MessageBarType, PrimaryButton, SearchBox } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { ILibraryResult, IScanResult } from '../models/IScanResult';
import { formatBytes } from '../services/formatBytes';
import {
  buildMap,
  coverage,
  isLarge,
  libraryFiles,
  libraryKey,
  librarySize,
  selectionTotals
} from '../services/siteMap';
import { formatCompact, formatDate } from './dashboard/format';
import styles from './StoragePulse.module.scss';
import { format } from './text';

export interface IScanMapProps {
  result: IScanResult;
  selected: Set<string>;
  onChange: (next: Set<string>) => void;
  onScan: () => void;
  onRefresh: () => void;
  onClose: () => void;
  canScan: boolean;
}

const COLLAPSE_ABOVE_LIBRARIES = 30;

function copy(source: Set<string>): Set<string> {
  const next = new Set<string>();
  source.forEach((k) => next.add(k));
  return next;
}

function status(lib: ILibraryResult): { text: string; kind: 'todo' | 'done' | 'failed' } {
  if (lib.unscanned || lib.pending) {
    return { text: strings.MapNotScanned, kind: 'todo' };
  }
  if (lib.error) {
    return { text: strings.MapFailed, kind: 'failed' };
  }
  return {
    text: lib.scannedAt ? format(strings.MapScannedOn, { date: formatDate(new Date(lib.scannedAt)) }) : strings.MapScannedOn.replace(' {date}', ''),
    kind: 'done'
  };
}

/**
 * The site as a tree of sites, subsites and libraries with SharePoint's own size
 * and file count for each, where the owner ticks what to read now. Whatever is
 * left unticked stays on the map for a later visit.
 */
export const ScanMap: React.FC<IScanMapProps> = ({ result, selected, onChange, onScan, onRefresh, onClose, canScan }) => {
  const [filter, setFilter] = useState('');
  const map = useMemo(() => buildMap(result), [result]);
  const cover = useMemo(() => coverage(result), [result]);
  const totals = useMemo(() => selectionTotals(result, selected), [result, selected]);
  const [expanded, setExpanded] = useState<Set<string>>(
    () => new Set(result.libraries.length > COLLAPSE_ABOVE_LIBRARIES ? [] : result.webs.map((w) => w.url))
  );

  const needle = filter.trim().toLowerCase();
  const matches = (lib: ILibraryResult): boolean =>
    !needle || lib.title.toLowerCase().indexOf(needle) >= 0 || lib.webTitle.toLowerCase().indexOf(needle) >= 0;

  const setKeys = (libs: ILibraryResult[], on: boolean): void => {
    const next = copy(selected);
    libs.forEach((l) => (on ? next.add(libraryKey(l)) : next.delete(libraryKey(l))));
    onChange(next);
  };
  const toggleExpanded = (url: string): void => {
    const next = copy(expanded);
    if (next.has(url)) {
      next.delete(url);
    } else {
      next.add(url);
    }
    setExpanded(next);
  };

  const share = cover.totalBytes > 0 ? Math.round((cover.scannedBytes / cover.totalBytes) * 100) : undefined;

  return (
    <section className={styles.mapPanel} aria-label={strings.MapTitle}>
      <div className={styles.mapHead}>
        <div>
          <h3 className={styles.mapTitle}>{strings.MapTitle}</h3>
          <p className={styles.mapIntro}>{strings.MapIntro}</p>
        </div>
        <DefaultButton text={strings.MapClose} iconProps={{ iconName: 'Cancel' }} onClick={onClose} />
      </div>

      <div className={styles.mapCoverage}>
        <div className={styles.mapCoverageText}>
          <strong>{format(strings.MapCoverage, { done: cover.scanned, total: cover.total })}</strong>
          {share !== undefined && <span> · {format(strings.MapCoverageShare, { percent: `${share}%` })}</span>}
        </div>
        <div
          className={styles.track}
          role="progressbar"
          aria-valuemin={0}
          aria-valuemax={cover.total}
          aria-valuenow={cover.scanned}
        >
          <span className={styles.trackFill} style={{ width: `${cover.total ? (cover.scanned / cover.total) * 100 : 0}%` }} />
        </div>
      </div>

      <div className={styles.mapToolbar}>
        <SearchBox
          placeholder={strings.MapFilterPlaceholder}
          ariaLabel={strings.MapFilterPlaceholder}
          value={filter}
          onChange={(_, v) => setFilter(v || '')}
          className={styles.mapSearch}
        />
        <DefaultButton
          text={strings.MapSelectUnscanned}
          onClick={() => setKeys(result.libraries.filter((l) => l.unscanned || l.pending || l.error), true)}
          disabled={!canScan}
        />
        <DefaultButton text={strings.MapSelectNone} onClick={() => onChange(new Set())} disabled={selected.size === 0} />
        <DefaultButton
          text={expanded.size > 0 ? strings.MapCollapseAll : strings.MapExpandAll}
          onClick={() => setExpanded(expanded.size > 0 ? new Set() : new Set(result.webs.map((w) => w.url)))}
        />
        <DefaultButton text={strings.MapRefresh} iconProps={{ iconName: 'Refresh' }} onClick={onRefresh} disabled={!canScan} />
      </div>

      <div className={styles.mapTree} role="tree">
        <div className={`${styles.mapRow} ${styles.mapHeaderRow}`} aria-hidden="true">
          <span />
          <span />
          <span />
          <span className={styles.mapNum}>{strings.MapColFiles}</span>
          <span className={styles.mapNum}>{strings.MapColSize}</span>
        </div>
        {map.map((node) => {
          const visibleOwn = node.own.filter(matches);
          const visibleAll = node.all.filter(matches);
          if (needle && visibleAll.length === 0) {
            return null;
          }
          const open = needle ? true : expanded.has(node.web.url);
          const selectedCount = visibleAll.filter((l) => selected.has(libraryKey(l))).length;
          const scannedCount = node.all.filter((l) => !l.unscanned && !l.pending).length;
          const bytes = visibleAll.reduce((sum, l) => sum + librarySize(l), 0);
          const files = visibleAll.reduce((sum, l) => sum + libraryFiles(l), 0);
          return (
            <div key={node.web.url} role="treeitem" aria-expanded={open}>
              <div className={`${styles.mapRow} ${styles.mapWebRow}`} style={{ paddingLeft: 8 + node.depth * 20 }}>
                <button
                  type="button"
                  className={styles.mapToggle}
                  onClick={() => toggleExpanded(node.web.url)}
                  aria-label={open ? strings.MapCollapseAll : strings.MapExpandAll}
                >
                  <Icon iconName={open ? 'ChevronDown' : 'ChevronRight'} />
                </button>
                <Checkbox
                  ariaLabel={format(strings.MapSiteAria, { name: node.web.title })}
                  checked={visibleAll.length > 0 && selectedCount === visibleAll.length}
                  indeterminate={selectedCount > 0 && selectedCount < visibleAll.length}
                  disabled={!canScan || visibleAll.length === 0}
                  onChange={(_, checked) => setKeys(visibleAll, !!checked)}
                />
                <span className={styles.mapName}>
                  <Icon iconName="Globe" className={styles.mapIcon} />
                  <strong>{node.web.title}</strong>
                  <span className={styles.mapMeta}>
                    {' '}
                    · {format(strings.MapSiteSummary, { scanned: scannedCount, total: node.all.length })}
                  </span>
                  {node.web.error && (
                    <span className={styles.mapWarn} title={node.web.error}>
                      {' '}
                      <Icon iconName="Warning" />
                    </span>
                  )}
                </span>
                <span className={styles.mapNum}>{formatCompact(files)}</span>
                <span className={styles.mapNum}>{formatBytes(bytes)}</span>
              </div>
              {open &&
                visibleOwn.map((lib) => {
                  const st = status(lib);
                  return (
                    <div
                      key={libraryKey(lib)}
                      role="treeitem"
                      className={`${styles.mapRow} ${styles.mapLibRow}`}
                      style={{ paddingLeft: 8 + (node.depth + 1) * 20 + 24 }}
                    >
                      <span />
                      <Checkbox
                        ariaLabel={format(strings.MapLibraryAria, { name: lib.title })}
                        checked={selected.has(libraryKey(lib))}
                        disabled={!canScan}
                        onChange={(_, checked) => setKeys([lib], !!checked)}
                      />
                      <span className={styles.mapName}>
                        <Icon iconName="DocLibrary" className={styles.mapIcon} />
                        {lib.title}{' '}
                        <span
                          className={`${styles.mapBadge} ${
                            st.kind === 'done' ? styles.mapBadgeDone : st.kind === 'failed' ? styles.mapBadgeFailed : ''
                          }`}
                        >
                          {st.text}
                        </span>
                        {isLarge(lib) && (
                          <span className={`${styles.mapBadge} ${styles.mapBadgeLarge}`} title={strings.MapLargeHint}>
                            {strings.MapLarge}
                          </span>
                        )}
                      </span>
                      <span className={styles.mapNum}>{formatCompact(libraryFiles(lib))}</span>
                      <span className={styles.mapNum}>{formatBytes(librarySize(lib))}</span>
                    </div>
                  );
                })}
            </div>
          );
        })}
        {map.length === 0 && <div className={styles.mapMeta}>{strings.MapNoLibraries}</div>}
      </div>
      <div className={styles.mapMeta}>{strings.MapSizeNote}</div>

      {totals.large > 0 && (
        <MessageBar messageBarType={MessageBarType.warning}>{format(strings.MapLargeWarning, { count: totals.large })}</MessageBar>
      )}
      <div className={styles.mapFooter}>
        <span aria-live="polite">
          {totals.count === 0
            ? strings.MapNoneSelected
            : format(strings.MapSelected, {
                count: totals.count,
                size: formatBytes(totals.bytes),
                files: formatCompact(totals.files)
              })}
        </span>
        {canScan && (
          <PrimaryButton
            text={format(strings.MapScanSelected, { count: totals.count })}
            iconProps={{ iconName: 'Search' }}
            disabled={totals.count === 0}
            onClick={onScan}
          />
        )}
      </div>
    </section>
  );
};
