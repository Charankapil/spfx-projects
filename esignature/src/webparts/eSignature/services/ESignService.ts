import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  ActivityType,
  IActivity,
  ICompletedPayload,
  IDelegatedPayload,
  IRequest,
  IRequestPayload,
  ISignedPayload,
  ITemplate
} from '../models/types';
import { DocumentService, baseName, fileName, folderOf, safeFileName } from './DocumentService';
import { newId, sha256Hex } from './hash';
import { emailHtml, NotificationService } from './NotificationService';
import { buildSignedPdf } from './PdfService';
import { Provisioning, STORE } from './Provisioning';
import { odataString, SpRest } from './SpRest';
import { computeStatus, endOfDay, IRequestStatus, ISlotState, reminderDue, sameEmail } from './status';

/**
 * Everything the UI does to a request goes through here: sending, signing,
 * declining, delegating, reminding, cancelling, finalising and verifying.
 * Each step writes an activity item first and only then sends mail, so the
 * audit trail is complete even when email delivery fails; mail problems come
 * back as warnings rather than errors.
 */

export interface IUser {
  name: string;
  email: string;
}

export interface IPersonSuggestion {
  name: string;
  email: string;
  detail: string;
}

export interface IRequestWithStatus {
  request: IRequest;
  activities: IActivity[];
  status: IRequestStatus;
}

export interface IActionResult {
  warnings: string[];
}

export interface ISendInput {
  payload: Omit<IRequestPayload, 'documentUrl' | 'documentHash' | 'senderName' | 'senderEmail' | 'sentOn' | 'v'>;
  pdf: ArrayBuffer;
}

interface ISpUser {
  EMail?: string;
  Title?: string;
  Name?: string;
}

interface ISpRequestItem {
  Id: number;
  RequestId: string;
  SenderStatus: string;
  Payload: string;
  Created: string;
  Author: ISpUser;
}

interface ISpActivityItem {
  Id: number;
  RequestId: string;
  EventType: ActivityType;
  Slot: string | null;
  DocHash: string | null;
  Payload: string | null;
  Created: string;
  Modified: string;
  Author: ISpUser;
}

const AUTHOR_SELECT = 'Author/EMail,Author/Title,Author/Name';

function emailOf(user: ISpUser | undefined): string {
  if (!user) {
    return '';
  }
  if (user.EMail) {
    return user.EMail;
  }
  // Claims login, e.g. "i:0#.f|membership|someone@contoso.com".
  const name = user.Name || '';
  return name.indexOf('|') >= 0 ? name.substring(name.lastIndexOf('|') + 1) : name;
}

function parse<T>(json: string | null | undefined, fallback: T): T {
  if (!json) {
    return fallback;
  }
  try {
    return JSON.parse(json) as T;
  } catch {
    return fallback;
  }
}

/**
 * The same adopted signature usually fills several fields; store each image
 * once and refer to it from the fields, so the activity stays well inside a
 * multi-line text column.
 */
function packValues(values: { [fieldId: string]: string }): { values: { [fieldId: string]: string }; images: string[] } {
  const images: string[] = [];
  const packed: { [fieldId: string]: string } = {};
  Object.keys(values).forEach(id => {
    const value = values[id];
    if (value.indexOf('data:image/') === 0) {
      let index = images.indexOf(value);
      if (index < 0) {
        index = images.push(value) - 1;
      }
      packed[id] = `image:${index}`;
    } else {
      packed[id] = value;
    }
  });
  return { values: packed, images };
}

function unpackValues(payload: { values?: { [fieldId: string]: string }; images?: string[] }): void {
  if (!payload || !payload.values || !payload.images) {
    return;
  }
  const { images, values } = payload;
  Object.keys(values).forEach(id => {
    const match = /^image:(\d+)$/.exec(values[id]);
    if (match) {
      values[id] = images[parseInt(match[1], 10)] || '';
    }
  });
  delete payload.images;
}

const MAX_PAYLOAD = 60000;

export const CONSENT_TEXT =
  'I agree to do business electronically and that my electronic signature is the legal equivalent of my handwritten signature on this document.';

export class ESignService {
  public readonly rest: SpRest;
  public readonly docs: DocumentService;
  public readonly provisioning: Provisioning;
  private mail: NotificationService;
  private webServerRelativeUrl: string;

  constructor(private context: WebPartContext, private pageUrl: string) {
    this.rest = new SpRest(context.spHttpClient, context.pageContext.web.absoluteUrl);
    this.docs = new DocumentService(this.rest, context.spHttpClient);
    this.webServerRelativeUrl = context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    this.provisioning = new Provisioning(this.rest, this.webServerRelativeUrl);
    this.mail = new NotificationService(context.msGraphClientFactory, this.rest);
  }

  public get me(): IUser {
    const user = this.context.pageContext.user;
    return { name: user.displayName, email: user.email || user.loginName };
  }

  public link(requestId: string): string {
    return `${this.pageUrl}?esignRequest=${encodeURIComponent(requestId)}`;
  }

  public get libraryUrl(): string {
    return `${this.webServerRelativeUrl}/${STORE.library}`;
  }

  // ---------------------------------------------------------------- reading

  private toRequest(item: ISpRequestItem): IRequest | undefined {
    const payload = parse<IRequestPayload | undefined>(item.Payload, undefined);
    if (!payload || !payload.recipients) {
      return undefined;
    }
    return {
      itemId: item.Id,
      requestId: item.RequestId,
      senderStatus: item.SenderStatus || 'Active',
      payload,
      payloadJson: item.Payload,
      authorEmail: emailOf(item.Author),
      authorName: item.Author?.Title || emailOf(item.Author),
      created: item.Created
    };
  }

  private toActivity(item: ISpActivityItem): IActivity {
    const payload = parse<unknown>(item.Payload, {});
    if (item.EventType === 'signed') {
      unpackValues(payload as { values?: { [fieldId: string]: string }; images?: string[] });
    }
    return {
      id: item.Id,
      requestId: item.RequestId,
      type: item.EventType,
      slot: item.Slot || undefined,
      authorEmail: emailOf(item.Author),
      authorName: item.Author?.Title || emailOf(item.Author),
      created: item.Created,
      modified: item.Modified,
      docHash: item.DocHash || undefined,
      payload
    };
  }

  /** The most recent requests on this site (newest first). */
  public async loadRequests(limit = 2000): Promise<IRequest[]> {
    const items = await this.rest.getAll<ISpRequestItem>(
      this.provisioning.listPath(
        STORE.requests,
        'items',
        `$select=Id,RequestId,SenderStatus,Payload,Created,${AUTHOR_SELECT}&$expand=Author&$orderby=Id desc&$top=500`
      ),
      limit
    );
    return items.map(i => this.toRequest(i)).filter((r): r is IRequest => !!r);
  }

  public async loadActivities(requestIds: string[]): Promise<IActivity[]> {
    const result: IActivity[] = [];
    // Keep each $filter well inside SharePoint's URL and expression limits.
    for (let i = 0; i < requestIds.length; i += 20) {
      const filter = requestIds
        .slice(i, i + 20)
        .map(id => `RequestId eq ${odataString(id)}`)
        .join(' or ');
      const items = await this.rest.getAll<ISpActivityItem>(
        this.provisioning.listPath(
          STORE.activity,
          'items',
          `$select=Id,RequestId,EventType,Slot,DocHash,Payload,Created,Modified,${AUTHOR_SELECT}&$expand=Author&$filter=${encodeURIComponent(filter)}&$top=500`
        )
      );
      result.push(...items.map(it => this.toActivity(it)));
    }
    return result;
  }

  public async loadAll(): Promise<IRequestWithStatus[]> {
    const requests = await this.loadRequests();
    const activities = await this.loadActivities(requests.map(r => r.requestId));
    const now = new Date();
    return requests.map(request => {
      const own = activities.filter(a => a.requestId === request.requestId);
      return { request, activities: own, status: computeStatus(request, own, now) };
    });
  }

  public async loadOne(requestId: string): Promise<IRequestWithStatus | undefined> {
    const items = await this.rest.getJson<{ value: ISpRequestItem[] }>(
      this.provisioning.listPath(
        STORE.requests,
        'items',
        `$select=Id,RequestId,SenderStatus,Payload,Created,${AUTHOR_SELECT}&$expand=Author&$filter=${encodeURIComponent(
          `RequestId eq ${odataString(requestId)}`
        )}`
      )
    );
    // If an item were ever duplicated, the earliest one is the real request.
    const request = items.value.sort((a, b) => a.Id - b.Id).map(i => this.toRequest(i))[0];
    if (!request) {
      return undefined;
    }
    const activities = await this.loadActivities([requestId]);
    return { request, activities, status: computeStatus(request, activities) };
  }

  // ---------------------------------------------------------------- writing

  private async addActivity(requestId: string, type: ActivityType, slot?: string, payload?: unknown, docHash?: string): Promise<void> {
    await this.rest.postJson(this.provisioning.listPath(STORE.activity, 'items'), {
      Title: type,
      RequestId: requestId,
      EventType: type,
      Slot: slot || null,
      DocHash: docHash || null,
      Payload: payload === undefined ? null : JSON.stringify(payload)
    });
  }

  private async notify(warnings: string[], to: string[], subject: string, body: Parameters<typeof emailHtml>[0], cc?: string[]): Promise<void> {
    try {
      await this.mail.send({ to, cc, subject, html: emailHtml(body) });
    } catch (e) {
      warnings.push((e as Error).message);
    }
  }

  public async send(input: ISendInput, onProgress: (step: string) => void): Promise<{ requestId: string; warnings: string[] }> {
    const warnings: string[] = [];
    const requestId = newId();
    const folder = `${this.libraryUrl}/${requestId}`;

    onProgress('Storing the document...');
    await this.docs.ensureFolder(folder);
    const pdfName = `${safeFileName(baseName(input.payload.documentName))}.pdf`;
    const documentUrl = await this.docs.upload(folder, pdfName, input.pdf);
    // Hash what SharePoint actually stored, so the fingerprint matches what signers download.
    const documentHash = await sha256Hex(await this.docs.readLocal(documentUrl));

    const me = this.me;
    const payload: IRequestPayload = {
      ...input.payload,
      v: 1,
      documentUrl,
      documentHash,
      senderName: me.name,
      senderEmail: me.email,
      sentOn: new Date().toISOString()
    };

    onProgress('Creating the request...');
    await this.rest.postJson(this.provisioning.listPath(STORE.requests, 'items'), {
      Title: payload.title,
      RequestId: requestId,
      SenderStatus: 'Active',
      Payload: JSON.stringify(payload)
    });
    await this.addActivity(requestId, 'sent', undefined, { recipients: payload.recipients.length }, documentHash);

    onProgress('Notifying recipients...');
    const created = await this.loadOne(requestId);
    if (created) {
      await this.notifyTurn(created, [], warnings);
      const ccs = created.status.ccs.map(c => c.email);
      if (ccs.length) {
        await this.notify(warnings, ccs, `For your information: ${payload.title}`, {
          heading: `${me.name} sent "${payload.title}" for signature`,
          lines: [`You have been copied on this request and will receive the signed document when everyone has signed.`, payload.message],
          buttonText: 'View request',
          link: this.link(requestId)
        });
      }
    }
    return { requestId, warnings };
  }

  /** Emails the signers whose turn it now is (skipping those already notified before). */
  private async notifyTurn(item: IRequestWithStatus, alreadyCurrent: string[], warnings: string[]): Promise<void> {
    const p = item.request.payload;
    for (const slot of item.status.current) {
      if (alreadyCurrent.indexOf(slot.original.id) >= 0) {
        continue;
      }
      await this.notify(warnings, [slot.current.email], `Signature requested: ${p.title}`, {
        heading: `${p.senderName} has requested your signature`,
        lines: [
          `Document: ${p.documentName}`,
          p.message,
          slot.current.privateMessage || '',
          p.expiresOn ? `Please sign by ${endOfDay(p.expiresOn).toLocaleDateString()}.` : ''
        ],
        buttonText: 'Review and sign',
        link: this.link(item.request.requestId),
        note: 'You will be asked to sign in with your work account. Do not forward this email.'
      });
    }
  }

  public async markViewed(item: IRequestWithStatus, slot: ISlotState): Promise<void> {
    if (!slot.viewedOn) {
      await this.addActivity(item.request.requestId, 'viewed', slot.original.id);
    }
  }

  /**
   * Records a signature. Before anything is written the stored document is
   * downloaded again and hashed: if it no longer matches the fingerprint
   * taken when the request was sent, signing is refused.
   */
  public async sign(item: IRequestWithStatus, slot: ISlotState, values: { [fieldId: string]: string }): Promise<IActionResult & { completed: boolean }> {
    const warnings: string[] = [];
    const request = item.request;
    const fresh = await this.loadOne(request.requestId);
    if (
      !fresh ||
      fresh.status.state !== 'InProgress' ||
      !fresh.status.current.some(s => s.original.id === slot.original.id && sameEmail(s.current.email, this.me.email))
    ) {
      throw new Error('This request changed while you had it open (it may have been cancelled, declined or already signed). Reload and try again.');
    }
    const documentHash = await sha256Hex(await this.docs.readLocal(request.payload.documentUrl));
    if (documentHash !== request.payload.documentHash) {
      throw new Error('The stored document no longer matches the one that was sent (its fingerprint changed). Signing has been blocked; contact the sender.');
    }
    const signed: ISignedPayload & { images: string[] } = {
      ...packValues(values),
      documentHash,
      requestHash: await sha256Hex(request.payloadJson),
      consent: CONSENT_TEXT,
      userAgent: navigator.userAgent,
      timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone || ''
    };
    if (JSON.stringify(signed).length > MAX_PAYLOAD) {
      throw new Error('Your signature images are too large to store. Adopt a simpler signature (for example a typed one) and try again.');
    }
    const before = fresh.status.current.map(s => s.original.id);
    await this.addActivity(request.requestId, 'signed', slot.original.id, signed, documentHash);

    const after = await this.loadOne(request.requestId);
    if (!after) {
      return { warnings, completed: false };
    }
    if (after.status.state === 'Completed') {
      warnings.push(...(await this.finalize(after)).warnings);
      return { warnings, completed: true };
    }
    await this.notifyTurn(after, before, warnings);
    await this.notify(warnings, [request.payload.senderEmail], `${slot.current.name} signed: ${request.payload.title}`, {
      heading: `${slot.current.name} signed "${request.payload.title}"`,
      lines: [`${after.status.signedCount} of ${after.status.signerCount} signers have signed.`],
      buttonText: 'View progress',
      link: this.link(request.requestId)
    });
    return { warnings, completed: false };
  }

  /**
   * Produces the signed PDF (stamped fields + certificate of completion),
   * saves it next to the original in the request folder, records its
   * fingerprint and emails everyone. Called by the last signer, or by the
   * sender from the request page if that failed.
   */
  public async finalize(item: IRequestWithStatus): Promise<IActionResult> {
    const warnings: string[] = [];
    const { request, status, activities } = item;
    if (status.state !== 'Completed' || status.completed) {
      return { warnings };
    }
    const original = await this.docs.readLocal(request.payload.documentUrl);
    const completedOn = new Date();
    const bytes = await buildSignedPdf(original, { request, status, activities, completedOn });
    const name = `${safeFileName(baseName(request.payload.documentName))} - signed.pdf`;
    const signedUrl = await this.docs.upload(folderOf(request.payload.documentUrl), name, bytes);
    const finalHash = await sha256Hex(await this.docs.readLocal(signedUrl));
    const completed: ICompletedPayload = { signedUrl, finalHash };
    await this.addActivity(request.requestId, 'completed', undefined, completed, finalHash);

    if (request.payload.saveToSource && request.payload.sourceUrl) {
      try {
        await this.copyToSource(request, bytes, name);
      } catch (e) {
        warnings.push(`The signed copy could not be saved next to the original (${(e as Error).message}). The sender can do this from the request page.`);
      }
    }

    const everyone = [request.payload.senderEmail].concat(status.slots.map(s => s.current.email));
    await this.notify(
      warnings,
      everyone,
      `Completed: ${request.payload.title}`,
      {
        heading: `"${request.payload.title}" has been signed by everyone`,
        lines: [`All ${status.signerCount} signer(s) have signed. The signed document, with its certificate of completion, is saved in SharePoint.`],
        buttonText: 'Open the signed document',
        link: this.link(request.requestId)
      },
      status.ccs.map(c => c.email)
    );
    return { warnings };
  }

  public async copyToSource(request: IRequest, bytes?: Uint8Array | ArrayBuffer, name?: string): Promise<void> {
    const source = request.payload.sourceUrl;
    if (!source) {
      throw new Error('This request has no source document location.');
    }
    let content = bytes;
    let target = name || fileName(source);
    if (!content) {
      const loaded = await this.loadOne(request.requestId);
      const done = loaded?.status.completed?.payload as ICompletedPayload | undefined;
      if (!done) {
        throw new Error('The request has not been completed yet.');
      }
      content = await this.docs.readLocal(done.signedUrl);
      target = fileName(done.signedUrl);
    }
    const webUrl = await this.docs.webUrlFor(source);
    const url = await this.docs.upload(folderOf(source), target, content, webUrl);
    await this.addActivity(request.requestId, 'copied', undefined, { url });
  }

  public async decline(item: IRequestWithStatus, slot: ISlotState, reason: string): Promise<IActionResult> {
    const warnings: string[] = [];
    const p = item.request.payload;
    await this.addActivity(item.request.requestId, 'declined', slot.original.id, { reason });
    const others = item.status.slots.filter(s => s !== slot && s.status === 'signed').map(s => s.current.email);
    await this.notify(
      warnings,
      [p.senderEmail],
      `Declined: ${p.title}`,
      {
        heading: `${slot.current.name} declined to sign "${p.title}"`,
        lines: [reason ? `Reason: ${reason}` : 'No reason was given.', 'The request has been closed and no further signatures will be collected.'],
        buttonText: 'View request',
        link: this.link(item.request.requestId)
      },
      others
    );
    return { warnings };
  }

  public async delegate(item: IRequestWithStatus, slot: ISlotState, to: IDelegatedPayload): Promise<IActionResult> {
    const warnings: string[] = [];
    if (sameEmail(to.toEmail, slot.current.email)) {
      throw new Error('Choose someone other than yourself.');
    }
    if (item.status.slots.some(s => s !== slot && sameEmail(s.current.email, to.toEmail))) {
      throw new Error(`${to.toName || to.toEmail} is already a signer on this request.`);
    }
    await this.addActivity(item.request.requestId, 'delegated', slot.original.id, to);
    const p = item.request.payload;
    await this.notify(warnings, [to.toEmail], `Signature requested: ${p.title}`, {
      heading: `${slot.current.name} has asked you to sign "${p.title}" on their behalf`,
      lines: [`Original request from ${p.senderName}.`, to.reason ? `Note: ${to.reason}` : '', p.message],
      buttonText: 'Review and sign',
      link: this.link(item.request.requestId)
    });
    await this.notify(warnings, [p.senderEmail], `Delegated: ${p.title}`, {
      heading: `${slot.current.name} delegated signing to ${to.toName || to.toEmail}`,
      lines: [to.reason ? `Reason: ${to.reason}` : ''],
      buttonText: 'View request',
      link: this.link(item.request.requestId)
    });
    return { warnings };
  }

  public async remind(item: IRequestWithStatus, automatic = false): Promise<IActionResult> {
    const warnings: string[] = [];
    const p = item.request.payload;
    await this.addActivity(item.request.requestId, 'reminded', undefined, { automatic });
    for (const slot of item.status.current) {
      await this.notify(warnings, [slot.current.email], `Reminder: please sign ${p.title}`, {
        heading: `Reminder: ${p.senderName} is waiting for your signature`,
        lines: [
          `Document: ${p.documentName}`,
          p.expiresOn ? `This request expires on ${endOfDay(p.expiresOn).toLocaleDateString()}.` : ''
        ],
        buttonText: 'Review and sign',
        link: this.link(item.request.requestId)
      });
    }
    return { warnings };
  }

  /** Sends any reminders that are due on the current user's own requests. */
  public async sendDueReminders(items: IRequestWithStatus[]): Promise<number> {
    let sent = 0;
    for (const item of items) {
      if (sameEmail(item.request.authorEmail, this.me.email) && reminderDue(item.request, item.status)) {
        await this.remind(item, true);
        sent++;
      }
    }
    return sent;
  }

  public async cancel(item: IRequestWithStatus, reason: string): Promise<IActionResult> {
    const warnings: string[] = [];
    const p = item.request.payload;
    await this.addActivity(item.request.requestId, 'cancelled', undefined, { reason });
    await this.rest.merge(this.provisioning.listPath(STORE.requests, `items(${item.request.itemId})`), { SenderStatus: 'Cancelled' });
    const people = item.status.slots.filter(s => s.status === 'signed' || s.status === 'pending').map(s => s.current.email);
    if (people.length) {
      await this.notify(warnings, people, `Cancelled: ${p.title}`, {
        heading: `${p.senderName} cancelled "${p.title}"`,
        lines: [reason ? `Reason: ${reason}` : '', 'No further action is needed from you.']
      });
    }
    return { warnings };
  }

  // -------------------------------------------------------------- templates

  public async templates(): Promise<ITemplate[]> {
    const items = await this.rest.getAll<{ Id: number; Title: string; Payload: string }>(
      this.provisioning.listPath(STORE.templates, 'items', '$select=Id,Title,Payload&$orderby=Title&$top=500')
    );
    const templates: ITemplate[] = [];
    items.forEach(i => {
      const template = parse<ITemplate | undefined>(i.Payload, undefined);
      if (template && template.recipients) {
        templates.push({ ...template, itemId: i.Id, name: i.Title });
      }
    });
    return templates;
  }

  public async saveTemplate(template: ITemplate): Promise<void> {
    const body = { Title: template.name, Payload: JSON.stringify({ ...template, itemId: undefined }) };
    if (template.itemId) {
      await this.rest.merge(this.provisioning.listPath(STORE.templates, `items(${template.itemId})`), body);
    } else {
      await this.rest.postJson(this.provisioning.listPath(STORE.templates, 'items'), body);
    }
  }

  public async deleteTemplate(itemId: number): Promise<void> {
    await this.rest.fetch(this.provisioning.listPath(STORE.templates, `items(${itemId})`), 'POST', {
      headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
    });
  }

  // ----------------------------------------------------------------- verify

  /** Finds activities whose recorded fingerprint matches a file the user dropped in. */
  public async findByHash(hash: string): Promise<IRequestWithStatus[]> {
    const items = await this.rest.getJson<{ value: ISpActivityItem[] }>(
      this.provisioning.listPath(
        STORE.activity,
        'items',
        `$select=Id,RequestId&$filter=${encodeURIComponent(`DocHash eq ${odataString(hash)}`)}&$top=50`
      )
    );
    const ids = items.value.map(i => i.RequestId).filter((id, i, all) => all.indexOf(id) === i);
    const found: IRequestWithStatus[] = [];
    for (const id of ids) {
      const item = await this.loadOne(id);
      if (item) {
        found.push(item);
      }
    }
    return found;
  }

  // ----------------------------------------------------------------- people

  public async searchPeople(query: string): Promise<IPersonSuggestion[]> {
    if (query.trim().length < 2) {
      return [];
    }
    const result = await this.rest.postJson<{ value: string }>(
      '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser',
      {
        queryParams: {
          QueryString: query,
          MaximumEntitySuggestions: 10,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalSource: 15,
          PrincipalType: 1,
          AllowMultipleEntities: false
        }
      }
    );
    const entries = parse<{ DisplayText: string; Description: string; EntityData?: { Email?: string; Title?: string } }[]>(result.value, []);
    return entries
      .map(e => ({
        name: e.DisplayText,
        email: (e.EntityData && e.EntityData.Email) || (/@/.test(e.Description) ? e.Description : ''),
        detail: (e.EntityData && e.EntityData.Title) || e.Description
      }))
      .filter(p => !!p.email);
  }
}
