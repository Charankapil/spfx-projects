import { MSGraphClientFactory } from '@microsoft/sp-http';
import { SpRest } from './SpRest';

/**
 * Email notifications, sent as the signed-in user.
 *
 * First choice is Microsoft Graph /me/sendMail through SPFx's built-in Graph
 * client (the package requests the delegated Mail.Send permission, which a
 * tenant admin approves once in the SharePoint admin center - no app
 * registration). If that is not approved, SharePoint's SendEmail utility is
 * tried, which only reaches people inside the organisation. If both fail the
 * caller gets the error and the UI offers a copyable link instead, so a
 * request is never stuck because mail could not be sent.
 */

export interface IEmail {
  to: string[];
  cc?: string[];
  subject: string;
  /** Pre-built HTML body. */
  html: string;
}

export class NotificationService {
  constructor(private graph: MSGraphClientFactory, private rest: SpRest) {}

  public async send(email: IEmail): Promise<void> {
    const to = unique(email.to);
    const cc = unique(email.cc || []).filter(c => to.indexOf(c) < 0);
    if (!to.length) {
      return;
    }
    let graphError: unknown;
    try {
      const client = await this.graph.getClient('3');
      await client.api('/me/sendMail').post({
        message: {
          subject: email.subject,
          body: { contentType: 'HTML', content: email.html },
          toRecipients: to.map(address => ({ emailAddress: { address } })),
          ccRecipients: cc.map(address => ({ emailAddress: { address } }))
        },
        saveToSentItems: true
      });
      return;
    } catch (e) {
      graphError = e;
    }
    try {
      await this.rest.postJson('/_api/SP.Utilities.Utility.SendEmail', {
        properties: { To: to, CC: cc, Subject: email.subject, Body: email.html }
      });
    } catch (e) {
      throw new Error(
        `Email could not be sent (Graph: ${(graphError as Error)?.message || graphError}; SharePoint: ${(e as Error).message}).`
      );
    }
  }
}

function unique(addresses: string[]): string[] {
  const seen: { [key: string]: boolean } = {};
  return addresses
    .map(a => a.trim())
    .filter(a => {
      const key = a.toLowerCase();
      if (!a || seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}

function escape(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** A plain, client-safe HTML email with one call-to-action button. */
export function emailHtml(options: { heading: string; lines: string[]; buttonText?: string; link?: string; note?: string }): string {
  const paragraphs = options.lines
    .filter(l => !!l)
    .map(l => `<p style="margin:0 0 12px;font-size:14px;line-height:20px;color:#323130">${escape(l).replace(/\n/g, '<br/>')}</p>`)
    .join('');
  const button =
    options.link && options.buttonText
      ? `<p style="margin:20px 0"><a href="${escape(options.link)}" style="background:#0f6cbd;color:#ffffff;text-decoration:none;padding:10px 20px;border-radius:4px;font-weight:600;font-size:14px;display:inline-block">${escape(options.buttonText)}</a></p>`
      : '';
  const note = options.note ? `<p style="margin:16px 0 0;font-size:12px;color:#605e5c">${escape(options.note)}</p>` : '';
  return (
    `<div style="font-family:Segoe UI,Helvetica,Arial,sans-serif;max-width:560px;padding:8px">` +
    `<h2 style="font-size:20px;font-weight:600;color:#242424;margin:0 0 16px">${escape(options.heading)}</h2>` +
    `${paragraphs}${button}${note}</div>`
  );
}
