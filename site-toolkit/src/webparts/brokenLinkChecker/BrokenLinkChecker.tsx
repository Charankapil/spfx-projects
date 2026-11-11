import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { DefaultButton, Dropdown, IDropdownOption, PrimaryButton, ProgressIndicator } from '@fluentui/react';

import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { downloadCsv } from '../../common/csv';
import { formatDateTime, formatNumber } from '../../common/format';
import { CancelToken, enqueue, errorText, isCancelled } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { classify, extractHrefs, LinkStatus, normalizeLink } from './links';

export interface IBrokenLinkCheckerProps {
  title: string;
  maxLinks: string;
}

interface IPageRow {
  Id: number;
  Title?: string;
  FileRef: string;
  FileLeafRef: string;
  CanvasContent1?: string;
}

interface ILinkResult {
  url: string;
  status: LinkStatus;
  code: number;
  pages: string[];
}

interface ISaved {
  checkedAt: string;
  pagesRead: number;
  external: number;
  skipped: number;
  results: ILinkResult[];
}

const MAX_PAGES = 200;
const PAGE_BATCH = 25;

const STATUS_TEXT: Record<LinkStatus, string> = {
  broken: 'Broken',
  restricted: 'No access for you',
  unknown: 'Could not be checked',
  ok: 'Working'
};

const FILTERS: IDropdownOption[] = [
  { key: 'problems', text: 'Broken links only' },
  { key: 'all', text: 'All checked links' }
];

export const BrokenLinkChecker: React.FC<IToolkitComponentProps<IBrokenLinkCheckerProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const maxLinks = parseInt(props.properties.maxLinks, 10) || 300;
  const key = cacheKey(kit.scope, 'links');

  const [saved, setSaved] = useState<ISaved | undefined>(undefined);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<{ label: string; done: number; total: number }>({ label: '', done: 0, total: 0 });
  const [error, setError] = useState<string | undefined>(undefined);
  const [filter, setFilter] = useState<string>('problems');
  const tokenRef = useRef<CancelToken | undefined>(undefined);

  useEffect(() => {
    setSaved(cacheGet<ISaved>('session', key));
  }, [key]);

  useEffect(
    () => () => {
      if (tokenRef.current) {
        tokenRef.current.cancel();
      }
    },
    []
  );

  const run = useCallback(() => {
    const token = new CancelToken();
    tokenRef.current = token;
    setRunning(true);
    setError(undefined);
    setSaved(undefined);
    setProgress({ label: 'Reading pages', done: 0, total: 0 });

    const work = async (): Promise<ISaved> => {
      const libraries = await kit.sp.getJson<{ value: Array<{ Id: string }> }>(
        `${kit.sp.webUrl}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id`,
        { token }
      );
      const pagesLibrary = (libraries.value || [])[0];
      if (!pagesLibrary) {
        throw new Error('This site has no Site Pages library.');
      }

      // Step 1: read the pages, a batch at a time, and collect their links.
      const foundOn: Record<string, string[]> = {};
      let external = 0;
      let pagesRead = 0;
      let next: string | undefined =
        `${kit.sp.webUrl}/_api/web/lists(guid'${pagesLibrary.Id}')/items` +
        `?$select=Id,Title,FileRef,FileLeafRef,CanvasContent1&$filter=FSObjType eq 0&$top=${PAGE_BATCH}`;
      while (next && pagesRead < MAX_PAGES) {
        const page: { items: IPageRow[]; next: string | undefined } = await kit.sp.getPage<IPageRow>(next, { token });
        for (const row of page.items) {
          pagesRead++;
          const pageUrl = `${kit.sp.origin}${row.FileRef}`;
          const pageName = row.Title || row.FileLeafRef;
          const seenHere: Record<string, boolean> = {};
          for (const href of extractHrefs(row.CanvasContent1 || '')) {
            const link = normalizeLink(href, pageUrl, kit.sp.origin);
            if (!link || seenHere[link.url]) {
              continue;
            }
            seenHere[link.url] = true;
            if (!link.internal) {
              external++;
              continue;
            }
            foundOn[link.url] = foundOn[link.url] || [];
            foundOn[link.url].push(pageName);
          }
        }
        next = page.next;
        if (mounted.current) {
          setProgress({ label: 'Reading pages', done: pagesRead, total: 0 });
        }
      }

      // Step 2: check each distinct link once, one request at a time through the page queue.
      const all = Object.keys(foundOn);
      const urls = all.slice(0, maxLinks);
      const results: ILinkResult[] = [];
      for (let i = 0; i < urls.length; i++) {
        const url = urls[i];
        let code = 0;
        try {
          const response = await enqueue(() => fetch(url, { method: 'HEAD', credentials: 'same-origin', redirect: 'follow' }), token);
          code = response.status;
        } catch (err) {
          if (isCancelled(err)) {
            throw err;
          }
          // Network failure or throttled out: reported as "could not be checked", never as broken.
          code = 0;
        }
        results.push({ url, status: classify(code), code, pages: foundOn[url] });
        if (mounted.current) {
          setProgress({ label: 'Checking links', done: i + 1, total: urls.length });
        }
      }
      return { checkedAt: new Date().toISOString(), pagesRead, external, skipped: all.length - urls.length, results };
    };

    work()
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        const order: Record<LinkStatus, number> = { broken: 0, unknown: 1, restricted: 2, ok: 3 };
        result.results.sort((a, b) => order[a.status] - order[b.status] || a.url.localeCompare(b.url));
        cacheSet('session', key, result, 60);
        setSaved(result);
        setRunning(false);
      })
      .catch((err) => {
        if (!mounted.current) {
          return;
        }
        setRunning(false);
        if (!isCancelled(err)) {
          setError(`The check stopped: ${errorText(err)}`);
        }
      });
  }, [key, kit, maxLinks, mounted]);

  const exportCsv = useCallback(() => {
    downloadCsv(
      'link-check',
      ['Status', 'HTTP code', 'Link', 'Found on pages'],
      (saved ? saved.results : []).map((r) => [STATUS_TEXT[r.status], r.code || '', r.url, r.pages.join('; ')])
    );
  }, [saved]);

  const results = saved ? saved.results : [];
  const broken = results.filter((r) => r.status === 'broken');
  const visible = filter === 'all' ? results : broken;

  return (
    <Shell
      icon="RemoveLink"
      title={props.properties.title || 'Broken Link Checker'}
      subtitle="Checks the links on this site's pages that point inside your SharePoint."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        <>
          {running ? (
            <DefaultButton iconProps={{ iconName: 'Stop' }} text="Stop" onClick={() => tokenRef.current && tokenRef.current.cancel()} />
          ) : (
            <PrimaryButton iconProps={{ iconName: 'Play' }} text={saved ? 'Check again' : 'Check links'} onClick={run} />
          )}
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={running || results.length === 0} onClick={exportCsv} />
        </>
      }
    >
      {running && (
        <ProgressIndicator
          label={progress.total > 0 ? `${progress.label}: ${progress.done} of ${progress.total}` : `${progress.label}: ${progress.done}`}
          description="Links are checked one at a time so the check stays gentle on SharePoint."
          percentComplete={progress.total > 0 ? progress.done / progress.total : undefined}
        />
      )}

      {!running && !saved && <Empty icon="RemoveLink" text="Press Check links to scan the pages of this site. Nothing is changed." />}

      {saved && (
        <>
          <div className={styles.kpis}>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatNumber(broken.length)}</div>
              <div className={styles.kpiLabel}>Broken links</div>
            </div>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatNumber(results.length)}</div>
              <div className={styles.kpiLabel}>Links checked</div>
            </div>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatNumber(saved.pagesRead)}</div>
              <div className={styles.kpiLabel}>Pages read</div>
            </div>
          </div>
          <div className={styles.small}>
            Checked {formatDateTime(saved.checkedAt)}. {formatNumber(saved.external)} link{saved.external === 1 ? '' : 's'} to other websites
            {saved.external === 1 ? ' was' : ' were'} not checked, because a browser is not allowed to test other sites.
            {saved.skipped > 0 ? ` ${formatNumber(saved.skipped)} more links were over the limit of ${maxLinks} per run.` : ''}
            {saved.pagesRead >= MAX_PAGES ? ` Only the first ${MAX_PAGES} pages are read.` : ''}
          </div>

          <div className={styles.toolbar}>
            <Dropdown
              ariaLabel="Show"
              options={FILTERS}
              selectedKey={filter}
              onChange={(_, option) => option && setFilter(String(option.key))}
              styles={{ root: { minWidth: 200 } }}
            />
          </div>

          {visible.length === 0 && (
            <Empty icon="Completed" text={filter === 'all' ? 'No internal links were found on the pages.' : 'No broken links were found.'} />
          )}

          {visible.length > 0 && (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Status</th>
                    <th>Link</th>
                    <th>Found on</th>
                  </tr>
                </thead>
                <tbody>
                  {visible.slice(0, 500).map((result) => (
                    <tr key={result.url}>
                      <td className={styles.nowrap}>
                        <span
                          className={`${styles.pill} ${
                            result.status === 'broken' ? styles.pillBad : result.status === 'ok' ? styles.pillGood : styles.pillWarn
                          }`}
                        >
                          {STATUS_TEXT[result.status]}
                          {result.code ? ` (${result.code})` : ''}
                        </span>
                      </td>
                      <td>
                        <a className={styles.link} href={result.url} target="_blank" rel="noopener noreferrer" style={{ wordBreak: 'break-all' }}>
                          {result.url}
                        </a>
                      </td>
                      <td>{result.pages.join(', ')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </Shell>
  );
};
