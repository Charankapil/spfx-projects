import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';

import { IFileTypeStat } from '../../models/IFileTypeStat';
import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import styles from './Dashboard.module.scss';
import { categoryOf, IFileCategory, OTHER_CATEGORY, totalsByCategory } from './fileTypeCategories';
import { formatPercent } from './format';
import { squarify } from './squarify';

/** Past this many types the tail folds into one tile, like WinDirStat's long extension list. */
const MAX_TILES = 80;

interface ITypeEntry {
  stat: IFileTypeStat;
  category: IFileCategory;
  value: number;
}

type TileDatum = { kind: 'type'; entry: ITypeEntry } | { kind: 'more'; count: number; value: number };

interface IHover {
  x: number;
  y: number;
  datum: TileDatum;
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

/**
 * WinDirStat-style extension view: one tile per file type across the whole
 * site collection, sized by number of files and coloured by category.
 */
export const Treemap: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const width = useWidth(containerRef);
  const height = Math.round(Math.max(280, Math.min(460, width * 0.5)));
  const [hover, setHover] = useState<IHover | undefined>(undefined);

  const formatValue = (v: number): string => `${v.toLocaleString()} files`;

  const entries = useMemo(() => {
    const list: ITypeEntry[] = overview.totalFileTypeStats
      .map((stat: IFileTypeStat) => ({ stat, category: categoryOf(stat.extension), value: stat.count }))
      .filter((e) => e.value > 0);
    list.sort((a, b) => b.value - a.value);
    return list;
  }, [overview]);

  const total = useMemo(() => entries.reduce((sum, e) => sum + e.value, 0), [entries]);

  const cells = useMemo(() => {
    if (width <= 0) {
      return [];
    }
    const shown = entries.slice(0, MAX_TILES);
    const rest = entries.slice(MAX_TILES);
    const items: { value: number; data: TileDatum }[] = shown.map((entry) => ({
      value: entry.value,
      data: { kind: 'type', entry }
    }));
    if (rest.length > 0) {
      const restValue = rest.reduce((sum, e) => sum + e.value, 0);
      items.push({ value: restValue, data: { kind: 'more', count: rest.length, value: restValue } });
    }
    return squarify(items, { x: 0, y: 0, w: width, h: height });
  }, [entries, width, height]);

  const legend = useMemo(() => totalsByCategory(overview.totalFileTypeStats), [overview]);

  const track = (e: React.MouseEvent, datum: TileDatum): void => {
    const box = containerRef.current ? containerRef.current.getBoundingClientRect() : undefined;
    if (!box) {
      return;
    }
    setHover({ x: e.clientX - box.left, y: e.clientY - box.top, datum });
  };

  const caption = 'Tile area is the number of files of each type across the whole site collection.';

  const tooltip = (datum: TileDatum): JSX.Element => {
    if (datum.kind === 'more') {
      return (
        <>
          <div className={styles.tooltipTitle}>{datum.count} smaller file types</div>
          <div>
            {formatValue(datum.value)} ({formatPercent(datum.value, total)})
          </div>
        </>
      );
    }
    const { stat, category } = datum.entry;
    return (
      <>
        <div className={styles.tooltipTitle}>
          .{stat.extension} <span className={styles.tooltipMuted}>{category.label}</span>
        </div>
        <div>
          {stat.count.toLocaleString()} files ({formatPercent(stat.count, total)} of all files)
        </div>
      </>
    );
  };

  return (
    <div className={styles.panel}>
      <h3 className={styles.panelTitle}>Files by file type</h3>
      <p className={styles.panelCaption}>{caption}</p>

      <div
        ref={containerRef}
        className={styles.treemap}
        style={{ height }}
        role="img"
        aria-label={`Treemap of ${entries.length} file types sized by number of files. The file type list shows exact figures.`}
        onMouseLeave={() => setHover(undefined)}
      >
        {entries.length === 0 && <div className={styles.empty}>No files found in the scanned libraries.</div>}

        {cells.map((cell) => {
          const x = cell.rect.x + 1;
          const y = cell.rect.y + 1;
          const w = Math.max(0, cell.rect.w - 2);
          const h = Math.max(0, cell.rect.h - 2);
          const datum = cell.data;
          const category = datum.kind === 'type' ? datum.entry.category : OTHER_CATEGORY;
          const label = datum.kind === 'type' ? `.${datum.entry.stat.extension}` : `+${datum.count} more`;
          const value = datum.kind === 'type' ? datum.entry.value : datum.value;

          return (
            <div
              key={datum.kind === 'type' ? datum.entry.stat.extension : '__more'}
              className={styles.piece}
              style={{ left: x, top: y, width: w, height: h, background: category.color, color: category.ink }}
              onMouseMove={(e) => track(e, datum)}
            >
              {w >= 34 && h >= 18 && <div className={styles.typeTileName}>{label}</div>}
              {w >= 64 && h >= 36 && <div className={styles.typeTileValue}>{formatValue(value)}</div>}
            </div>
          );
        })}

        {hover && (
          <div
            className={styles.tooltip}
            style={{
              left: Math.min(hover.x + 14, Math.max(0, width - 290)),
              top: hover.y + 14 > height - 100 ? Math.max(0, hover.y - 110) : hover.y + 14
            }}
          >
            {tooltip(hover.datum)}
          </div>
        )}
      </div>

      <div className={styles.legend}>
        {legend.map((t) => (
          <span key={t.category.key} className={styles.legendItem}>
            <span className={styles.swatch} style={{ background: t.category.color }} />
            {t.category.label} {formatPercent(t.count, total)}
          </span>
        ))}
      </div>
    </div>
  );
};
