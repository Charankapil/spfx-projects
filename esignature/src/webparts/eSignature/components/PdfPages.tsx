import * as React from 'react';
import { Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import type { IPdfDocument } from 'pdfjs-dist/legacy/build/pdf';

import { openPdf } from '../services/PdfService';
import { errorText } from './shared';
import styles from './ESign.module.scss';

export interface IPageBox {
  index: number;
  width: number;
  height: number;
}

export interface IPdfPagesProps {
  bytes: ArrayBuffer;
  /** Rendered on top of each page; positions are in CSS pixels of the rendered page. */
  renderOverlay?: (page: IPageBox) => React.ReactNode;
  onPageClick?: (page: IPageBox, x: number, y: number) => void;
  onLoaded?: (pageCount: number) => void;
  pageClassName?: string;
  /** Scroll this page into view when it changes. */
  scrollToPage?: number;
}

interface IPageSize {
  width: number;
  height: number;
}

/**
 * Renders every page of a PDF to a canvas sized to the container, with an
 * overlay layer per page for fields. Pages render one after another so a
 * long document shows its first page quickly.
 */
export function PdfPages(props: IPdfPagesProps): JSX.Element {
  const container = React.useRef<HTMLDivElement>(null);
  const [doc, setDoc] = React.useState<IPdfDocument>();
  const [sizes, setSizes] = React.useState<IPageSize[]>([]);
  const [width, setWidth] = React.useState(0);
  const [error, setError] = React.useState<string>();
  const canvases = React.useRef<(HTMLCanvasElement | null)[]>([]);
  const pageRefs = React.useRef<(HTMLDivElement | null)[]>([]);
  // pdf.js refuses two renders into one canvas at once, so a resize waits for the previous pass.
  const renderQueue = React.useRef<Promise<void>>(Promise.resolve());

  React.useEffect(() => {
    let disposed = false;
    let opened: IPdfDocument | undefined;
    setDoc(undefined);
    setSizes([]);
    setError(undefined);
    (async () => {
      try {
        opened = await openPdf(props.bytes);
        const list: IPageSize[] = [];
        for (let i = 1; i <= opened.numPages; i++) {
          const viewport = (await opened.getPage(i)).getViewport({ scale: 1 });
          list.push({ width: viewport.width, height: viewport.height });
        }
        if (!disposed) {
          setSizes(list);
          setDoc(opened);
          props.onLoaded?.(opened.numPages);
        }
      } catch (e) {
        if (!disposed) {
          setError(`This PDF could not be opened: ${errorText(e)}`);
        }
      }
    })();
    return () => {
      disposed = true;
      opened?.destroy().catch(() => undefined);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.bytes]);

  React.useEffect(() => {
    const element = container.current;
    if (!element) {
      return;
    }
    const measure = (): void => setWidth(Math.min(element.clientWidth - 32, 900));
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  React.useEffect(() => {
    if (!doc || width <= 0) {
      return;
    }
    let cancelled = false;
    const pass = async (): Promise<void> => {
      const ratio = window.devicePixelRatio || 1;
      for (let i = 0; i < doc.numPages && !cancelled; i++) {
        const canvas = canvases.current[i];
        if (!canvas) {
          continue;
        }
        const page = await doc.getPage(i + 1);
        const scale = (width / sizes[i].width) * ratio;
        const viewport = page.getViewport({ scale });
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const context = canvas.getContext('2d');
        if (context) {
          await page.render({ canvasContext: context, viewport }).promise;
        }
        page.cleanup();
      }
    };
    renderQueue.current = renderQueue.current.then(pass).catch(e => {
      if (!cancelled) {
        setError(errorText(e));
      }
    });
    return () => {
      cancelled = true;
    };
  }, [doc, width, sizes]);

  React.useEffect(() => {
    if (props.scrollToPage !== undefined) {
      pageRefs.current[props.scrollToPage]?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, [props.scrollToPage]);

  return (
    <div className={styles.pages} ref={container}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {!doc && !error && <Spinner label="Opening document..." />}
      {doc &&
        width > 0 &&
        sizes.map((size, index) => {
          const box: IPageBox = { index, width, height: (width * size.height) / size.width };
          return (
            <React.Fragment key={index}>
              <div
                ref={el => (pageRefs.current[index] = el)}
                className={`${styles.page} ${props.pageClassName || ''}`}
                style={{ width: box.width, height: box.height }}
                onMouseDown={e => {
                  if (props.onPageClick && e.target === e.currentTarget.firstChild) {
                    const rect = e.currentTarget.getBoundingClientRect();
                    props.onPageClick(box, e.clientX - rect.left, e.clientY - rect.top);
                  }
                }}
              >
                <canvas ref={el => (canvases.current[index] = el)} />
                {props.renderOverlay?.(box)}
              </div>
              <div className={styles.pageNumber}>
                Page {index + 1} of {sizes.length}
              </div>
            </React.Fragment>
          );
        })}
    </div>
  );
}
