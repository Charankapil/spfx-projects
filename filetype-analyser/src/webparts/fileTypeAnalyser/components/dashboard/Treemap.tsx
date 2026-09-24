import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';

import { ILibraryNode } from '../../models/ILibraryNode';
import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import { IWebNode } from '../../models/IWebNode';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { ICategoryTotal, totalsByCategory } from './fileTypeCategories';
import { formatPercent } from './format';
import { squarify } from './squarify';

/** Past this many libraries the tail folds into one tile, keeping the DOM small on huge sites. */
const MAX_TILES = 150;
const HEADER_HEIGHT = 22;

interface ILibraryLocation {
  library: ILibraryNode;
  webTitle: string;
  inSubsite: boolean;
}

interface ILibraryEntry extends ILibraryLocation {
  value: number;
}

type TileDatum = { kind: 'library'; entry: ILibraryEntry } | { kind: 'more'; count: number };

interface IHover {
  x: number;
  y: number;
  entry: ILibraryEntry;
  category?: ICategoryTotal;
}

function collect(web: IWebNode, out: ILibraryLocation[], inSubsite: boolean): void {
  web.libraries.forEach((library) => out.push({ library, webTitle: web.title, inSubsite }));
  web.webs.forEach((child) => collect(child, out, true));
}

function useWidth(ref: React.RefObject<HTMLDivElement>): number {
  const [width, setWidth] = useState(0);
  useEffect(() => {
    const el = ref.current;
    if (!el) {
      return undefined;
    }
    const update = (): void => setWidth(el.clientWidth);
    update();
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(update);
      observer.observe(el);
      return () => observer.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, [ref]);
  return width;
}

export const Treemap: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const width = useWidth(containerRef);
  const height = Math.round(Math.max(280, Math.min(460, width * 0.5)));
  const [hover, setHover] = useState<IHover | undefined>(undefined);

  const bySize = overview.sizesAvailable;

  const { entries, unsized } = useMemo(() => {
    const all: ILibraryLocation[] = [];
    collect(overview.rootWeb, all, false);
    const list: ILibraryEntry[] = [];
    let missing = 0;
    all.forEach((location) => {
      const { library } = location;
      if (bySize && typeof library.sizeBytes !== 'number') {
        missing++;
        return;
      }
      const value = bySize ? library.sizeBytes || 0 : library.totalFiles;
      if (value > 0) {
        list.push({ ...location, value });
      }
    });
    list.sort((a, b) => b.value - a.value);
    return { entries: list, unsized: missing };
  }, [overview, bySize]);

  const total = useMemo(() => entries.reduce((sum, e) => sum + e.value, 0), [entries]);

  const cells = useMemo(() => {
    if (width <= 0) {
      return [];
    }
    const shown = entries.slice(0, MAX_TILES);
    const rest = entries.slice(MAX_TILES);
    const items: { value: number; data: TileDatum }[] = shown.map((entry) => ({
      value: entry.value,
      data: { kind: 'library', entry }
    }));
    if (rest.length > 0) {
      items.push({
        value: rest.reduce((sum, e) => sum + e.value, 0),
        data: { kind: 'more', count: rest.length }
      });
    }
    return squarify(items, { x: 0, y: 0, w: width, h: height });
  }, [entries, width, height]);

  const legend = useMemo(() => totalsByCategory(overview.totalFileTypeStats), [overview]);
  const legendTotal = legend.reduce((sum, t) => sum + t.count, 0);

  const track = (e: React.MouseEvent, entry: ILibraryEntry, category?: ICategoryTotal): void => {
    const box = containerRef.current ? containerRef.current.getBoundingClientRect() : undefined;
    if (!box) {
      return;
    }
    setHover({ x: e.clientX - box.left, y: e.clientY - box.top, entry, category });
  };

  const caption = bySize
    ? 'Tile area is each library’s storage (including versions, from SharePoint storage metrics). ' +
      'Colours inside a tile show that library’s files by type — as a share of file count, not bytes.'
    : 'Storage figures were not available for this scan, so tile area is the number of files in each library. ' +
      'Colours inside a tile show its files by type.';

  return (
    <div className={styles.panel}>
      <h3 className={styles.panelTitle}>{bySize ? 'Storage by library' : 'Files by library'}</h3>
      <p className={styles.panelCaption}>
        {caption}
        {bySize && unsized > 0 && ` ${unsized} ${unsized === 1 ? 'library has' : 'libraries have'} no size and ${unsized === 1 ? 'is' : 'are'} not shown.`}
      </p>

      <div
        ref={containerRef}
        className={styles.treemap}
        style={{ height }}
        role="img"
        aria-label={`Treemap of ${entries.length} libraries sized by ${bySize ? 'storage' : 'file count'}. The file type list shows exact figures.`}
        onMouseLeave={() => setHover(undefined)}
      >
        {entries.length === 0 && <div className={styles.empty}>No files found in the scanned libraries.</div>}

        {cells.map((cell, i) => {
          const x = cell.rect.x + 1;
          const y = cell.rect.y + 1;
          const w = Math.max(0, cell.rect.w - 2);
          const h = Math.max(0, cell.rect.h - 2);
          const datum = cell.data;

          if (datum.kind === 'more') {
            return (
              <div key="more" className={`${styles.tile} ${styles.moreTile}`} style={{ left: x, top: y, width: w, height: h }}>
                {w > 50 && h > 24 ? `+${datum.count} more libraries` : ''}
              </div>
            );
          }

          const { entry } = datum;
          const showHeader = w >= 72 && h >= 48;
          const bodyTop = showHeader ? HEADER_HEIGHT : 0;
          const categories = totalsByCategory(entry.library.fileTypes);
          const pieces = squarify(
            categories.map((c) => ({ value: c.count, data: c })),
            { x: 0, y: bodyTop, w, h: h - bodyTop }
          );

          return (
            <div
              key={`${entry.library.id}-${i}`}
              className={styles.tile}
              style={{ left: x, top: y, width: w, height: h }}
              onMouseMove={(e) => track(e, entry)}
            >
              {showHeader && (
                <div className={styles.tileHeader}>
                  <span className={styles.tileTitle}>
                    {/* Every subsite has its own "Documents"; the site name tells them apart. */}
                    {entry.inSubsite ? `${entry.library.title} · ${entry.webTitle}` : entry.library.title}
                  </span>
                  <span className={styles.tileSize}>
                    {bySize ? formatBytes(entry.value) : `${entry.value.toLocaleString()} files`}
                  </span>
                </div>
              )}
              {pieces.map((p) => {
                const pw = Math.max(0, p.rect.w - 2);
                const ph = Math.max(0, p.rect.h - 2);
                const cat = p.data.category;
                return (
                  <div
                    key={cat.key}
                    className={styles.piece}
                    style={{
                      left: p.rect.x + 1,
                      top: p.rect.y + 1,
                      width: pw,
                      height: ph,
                      background: cat.color,
                      color: cat.ink
                    }}
                    onMouseMove={(e) => {
                      e.stopPropagation();
                      track(e, entry, p.data);
                    }}
                  >
                    {pw >= 48 && ph >= 20 ? cat.label : ''}
                  </div>
                );
              })}
            </div>
          );
        })}

        {hover && (
          <div
            className={styles.tooltip}
            style={{
              left: Math.min(hover.x + 14, Math.max(0, width - 290)),
              top: hover.y + 14 > height - 110 ? Math.max(0, hover.y - 120) : hover.y + 14
            }}
          >
            <div className={styles.tooltipTitle}>{hover.entry.library.title}</div>
            <div className={styles.tooltipMuted}>{hover.entry.webTitle}</div>
            {bySize && (
              <div>
                Storage: {formatBytes(hover.entry.value)} ({formatPercent(hover.entry.value, total)} of shown)
              </div>
            )}
            <div>Files: {hover.entry.library.totalFiles.toLocaleString()}</div>
            {hover.category && (
              <div>
                {hover.category.category.label}: {hover.category.count.toLocaleString()} files (
                {formatPercent(hover.category.count, hover.entry.library.totalFiles)} of this library)
                <div className={styles.tooltipMuted}>
                  {hover.category.extensions
                    .slice(0, 4)
                    .map((ext) => `.${ext.extension} ${ext.count.toLocaleString()}`)
                    .join(' · ')}
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      <div className={styles.legend}>
        {legend.map((t) => (
          <span key={t.category.key} className={styles.legendItem}>
            <span className={styles.swatch} style={{ background: t.category.color }} />
            {t.category.label} {formatPercent(t.count, legendTotal)}
          </span>
        ))}
      </div>
    </div>
  );
};
