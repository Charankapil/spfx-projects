import { IUniqueObject, ObjectKind } from '../models/IUniqueObject';
import { errorMessageFromBody, IBatchRequest, IBatchResponse } from './batch';
import { runPool } from './pool';
import { HttpError, isCancelled, SpRest } from './SpRest';

/** SharePoint accepts up to 100 requests per $batch. */
const BATCH_SIZE = 100;
const MAX_INNER_RETRIES = 4;

const KIND_ORDER: { [kind in ObjectKind]: number } = {
  web: 0,
  list: 1,
  library: 1,
  folder: 2,
  file: 3,
  item: 3
};

const ROLE_ASSIGNMENTS_QUERY =
  'roleassignments?$expand=Member,RoleDefinitionBindings' +
  '&$select=Member/Title,Member/LoginName,RoleDefinitionBindings/Name';

type Collection<T> = T[] | { results?: T[] } | undefined;

interface IRoleAssignment {
  Member?: { Title?: string; LoginName?: string };
  RoleDefinitionBindings?: Collection<{ Name: string }>;
}

function toArray<T>(collection: Collection<T>): T[] {
  if (!collection) {
    return [];
  }
  return Array.isArray(collection) ? collection : collection.results || [];
}

/**
 * "Finance Members: Contribute; Jane Doe: Read, Edit". Limited Access is left
 * out: SharePoint grants it automatically and it comes back on its own.
 */
export function formatRoleAssignments(body: string): string {
  const json = JSON.parse(body);
  const assignments = toArray<IRoleAssignment>(json.value || json.d?.results || json.d);
  return assignments
    .map((ra) => {
      const who = ra.Member?.Title || ra.Member?.LoginName || 'Unknown principal';
      const roles = toArray(ra.RoleDefinitionBindings)
        .map((r) => r.Name)
        .filter((n) => n && n !== 'Limited Access');
      return roles.length > 0 ? `${who}: ${roles.join(', ')}` : '';
    })
    .filter((s) => s.length > 0)
    .join('; ');
}

export interface IRestoreProgress {
  done: number;
  total: number;
  restored: number;
  failed: number;
  current: string;
}

function objectApi(o: IUniqueObject): string {
  const web = `${o.webUrl}/_api/web`;
  if (o.kind === 'web') {
    return web;
  }
  const list = `${web}/lists(guid'${o.listId}')`;
  return o.kind === 'list' || o.kind === 'library' ? list : `${list}/items(${o.itemId})`;
}

function isThrottled(r: IBatchResponse): boolean {
  return r.status === 429 || r.status === 503;
}

function isSuccess(r: IBatchResponse): boolean {
  return r.status >= 200 && r.status < 300;
}

/**
 * Resets role inheritance on the given objects, 100 per $batch round trip.
 * With backup on, each object's current role assignments are read first (in
 * a batch of GETs) and kept in the report; an object whose permissions could
 * not be read is left alone rather than reset without a record.
 */
export class InheritanceRestorer {
  constructor(
    private rest: SpRest,
    private onProgress: (progress: IRestoreProgress) => void,
    /** Batches in flight at once: 1 in gentle mode, 2 in standard. */
    private parallelBatches = 1
  ) {}

  public async restore(targets: IUniqueObject[], backup: boolean): Promise<void> {
    // Sites, then lists, then folders from the top down, then files. The
    // order is not needed for correctness but makes a report cut short by
    // cancelling read naturally.
    const ordered = [...targets].sort(
      (a, b) => KIND_ORDER[a.kind] - KIND_ORDER[b.kind] || a.depth - b.depth
    );

    const byWeb = new Map<string, IUniqueObject[]>();
    for (const o of ordered) {
      byWeb.set(o.webUrl, [...(byWeb.get(o.webUrl) || []), o]);
    }
    const chunks: IUniqueObject[][] = [];
    byWeb.forEach((objects) => {
      for (let i = 0; i < objects.length; i += BATCH_SIZE) {
        chunks.push(objects.slice(i, i + BATCH_SIZE));
      }
    });

    const progress: IRestoreProgress = { done: 0, total: ordered.length, restored: 0, failed: 0, current: '' };
    this.onProgress({ ...progress });
    await runPool(chunks, this.parallelBatches, async (chunk) => {
      progress.current = chunk[0].path;
      await this.processChunk(chunk, backup);
      progress.done += chunk.length;
      for (const o of chunk) {
        if (o.status === 'restored') {
          progress.restored++;
        } else if (o.status === 'failed') {
          progress.failed++;
        }
      }
      this.onProgress({ ...progress });
    });
  }

  private async processChunk(chunk: IUniqueObject[], backup: boolean): Promise<void> {
    const webUrl = chunk[0].webUrl;
    let toReset = chunk;

    if (backup) {
      const reads = await this.run(
        webUrl,
        chunk.map((o) => ({ method: 'GET', url: `${objectApi(o)}/${ROLE_ASSIGNMENTS_QUERY}` }))
      );
      toReset = [];
      reads.forEach((r, i) => {
        const o = chunk[i];
        if (isSuccess(r)) {
          try {
            o.previousPermissions = formatRoleAssignments(r.body);
            toReset.push(o);
            return;
          } catch {
            // Unreadable body: treat as a failed backup below.
          }
        }
        o.status = 'failed';
        o.message = `Not restored: its current permissions could not be backed up (${errorMessageFromBody(
          r.body,
          r.statusText || String(r.status)
        )}).`;
      });
    }
    if (toReset.length === 0) {
      return;
    }

    const results = await this.run(
      webUrl,
      toReset.map((o) => ({ method: 'POST', url: `${objectApi(o)}/ResetRoleInheritance` }))
    );
    const now = new Date().toISOString();
    results.forEach((r, i) => {
      const o = toReset[i];
      if (isSuccess(r)) {
        o.status = 'restored';
        o.restoredAt = now;
        o.message = undefined;
      } else {
        o.status = 'failed';
        o.message = errorMessageFromBody(r.body, r.statusText || `HTTP ${r.status}`);
      }
    });
  }

  /**
   * One $batch, falling back to one call at a time if the batch endpoint
   * itself fails. Requests throttled inside the batch are retried after a
   * pause, so throttling never turns into a "failed" row in the report.
   */
  private async run(webUrl: string, requests: IBatchRequest[]): Promise<IBatchResponse[]> {
    const results = await this.runOnce(webUrl, requests);
    for (let attempt = 1; attempt <= MAX_INNER_RETRIES; attempt++) {
      const throttled = results.map((r, i) => (isThrottled(r) ? i : -1)).filter((i) => i >= 0);
      if (throttled.length === 0) {
        break;
      }
      // Tell the client, so every other request backs off too, not just this batch.
      this.rest.noteThrottled(5 * attempt);
      await this.rest.sleep(5000 * attempt);
      const retried = await this.runOnce(
        webUrl,
        throttled.map((i) => requests[i])
      );
      throttled.forEach((index, k) => {
        results[index] = retried[k];
      });
    }
    return results;
  }

  private async runOnce(webUrl: string, requests: IBatchRequest[]): Promise<IBatchResponse[]> {
    try {
      return await this.rest.batch(webUrl, requests);
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      console.warn('[ReInherit] $batch failed, falling back to single requests.', err);
    }
    const results: IBatchResponse[] = [];
    for (const request of requests) {
      this.rest.throwIfCancelled();
      try {
        if (request.method === 'GET') {
          const json = await this.rest.getJson<unknown>(request.url);
          results.push({ status: 200, statusText: 'OK', body: JSON.stringify(json) });
        } else {
          const status = await this.rest.post(request.url, request.body);
          results.push({ status, statusText: '', body: '' });
        }
      } catch (err) {
        if (isCancelled(err)) {
          throw err;
        }
        results.push(
          err instanceof HttpError
            ? { status: err.status, statusText: err.detail, body: '' }
            : { status: 0, statusText: err instanceof Error ? err.message : 'Request failed', body: '' }
        );
      }
    }
    return results;
  }
}
