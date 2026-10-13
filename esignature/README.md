# eSignature

**Send documents for signature, sign them and track every request, without
leaving SharePoint.** One `.sppkg`, no Azure app registration, no
third-party signing service, no per-envelope fees.

eSignature is a SharePoint Framework solution with two parts:

- the **eSignature web part**: a dashboard, a send wizard, the signing
  screen, request details with the full audit trail, templates and a
  verification tool;
- a **Request signatures** button in every document library's command bar,
  which opens the send wizard with the selected PDF or Word file.

## Features

| Area | What you get |
| --- | --- |
| Documents | PDF and Word (`.docx`, `.doc`, `.docm`, `.dotx`, `.odt`, `.rtf`; converted to PDF by SharePoint). Pick from any library on the site, start from the library command bar, or upload from your computer. Password-protected PDFs are rejected up front. |
| Recipients | Any number of signers plus CC recipients. People search from the tenant directory, or type any email address. |
| Signing order | Parallel (everyone at once), sequential, or mixed: signers with the same step number sign in parallel and the next step opens when the previous one is complete. |
| Fields | Signature, initials, date signed, full name, email, free text, checkbox. Drag-and-drop placement on the rendered pages, resize, keyboard nudge, per-field required flag, colour-coded per signer. |
| Signing | Draw (speed-sensitive strokes), type in a handwriting font, or upload a scan (white background removed). Adopt once and it fills every signature or initials field; optionally remembered on the device. Guided "next required field", progress meter, and an explicit consent to sign electronically. |
| Decline and delegate | Signers can decline with a reason (closes the request) or assign their signature to someone else. |
| Sender controls | Cancel with a reason, send a reminder now, automatic reminders (daily, every 2 or 3 days, weekly), expiry date, private note per recipient, copy link. |
| Notifications | Emails for "your turn", progress, reminders, declines, delegations, cancellations and completion, sent as the user through Microsoft Graph (falls back to SharePoint's mail utility). |
| Completion | The signed PDF is generated with every value stamped in place and a **certificate of completion** appended: request details, each signer's identity, view and sign times, time zone, browser, consent text, signature, document fingerprints and the full event log. Saved to the request folder and, optionally, next to the original document. |
| Tamper evidence | SHA-256 of the document is taken at sending; every signer's browser re-downloads and re-hashes the document before signing and refuses to sign if it changed. The signed copy's SHA-256 is recorded too, and **Verify** checks any PDF against the audit trail without uploading it. |
| Audit | Timeline on every request, per-signer status, integrity panel, CSV export of the audit trail. |
| Templates | Save recipients (as roles), signing order, fields, message, reminders and expiry as a template, and reuse it with a new document. |
| Dashboard | Action required, waiting for others, completed, declined/cancelled/expired, all; search by title, document or person; progress per request. |
| Hosting | SharePoint pages and full-width pages, Teams tabs and personal apps. Picks up the site theme. |

## How it compares with Microsoft SharePoint eSignature

|                                          | SharePoint eSignature (Syntex)        | This solution |
| ---------------------------------------- | ------------------------------------- | ------------- |
| Cost                                     | Pay-as-you-go, per request            | None          |
| Signers per request                      | Limited (10 at the time of writing)   | No fixed limit |
| Mixed parallel + sequential steps        | -                                     | Yes |
| Delegate to someone else                 | -                                     | Yes |
| Reusable templates                       | -                                     | Yes |
| Verify a PDF against the audit trail     | -                                     | Yes |
| External signers without a guest account | Yes (email one-time passcode)         | No - see *Limitations* |
| Third-party providers (Adobe, DocuSign)  | Yes                                   | No |

Microsoft's feature set changes regularly; check its current documentation
before deciding between the two.

## How it works

All data stays on the site that hosts the web part, in four places created
by the Setup screen:

| Name | Holds | Who can edit an item |
| --- | --- | --- |
| eSignature Requests (`Lists/ESignRequests`) | One item per request: recipients, fields, settings, document fingerprint | Only the sender who created it |
| eSignature Activity (`Lists/ESignActivity`) | One item per event: sent, viewed, signed (with the signature and field values), declined, delegated, reminded, cancelled, completed | Only the person who created it |
| eSignature Templates (`Lists/ESignTemplates`) | Reusable templates | Anyone with edit rights |
| eSignature Documents (`ESignDocuments`) | A folder per request with the frozen PDF and the signed copy | Anyone with edit rights |

**The status of a request is never stored; it is calculated.** Every
screen recomputes it from the request and its activity items, using the
identity SharePoint itself stamped on each item (`Author`), not anything
written inside it:

- a signature counts only if it was created by the person currently holding
  that signer slot, and only once it is their turn;
- delegation hands the slot to someone else, and only the new holder's
  signature counts from then on;
- only the sender can cancel;
- an activity edited after it was recorded is ignored, flagged in the audit
  trail and in the certificate, because the list only lets people edit their
  own items and a signature must not change once given.

Combined with "people can only edit items they created" on both lists, this
means nobody can record a signature for someone else or alter another
person's actions through the UI, REST or the list itself. Ignored entries
stay visible in the history as **Not counted**, with the reason.

PDF work happens in the browser: [pdf.js](https://github.com/mozilla/pdf.js)
renders pages (its worker runs in-page, so nothing is loaded from a CDN and
tenant content security policy is not an issue) and
[pdf-lib](https://pdf-lib.js.org/) writes the signed PDF. Both are loaded
only when a document is opened.

## No Azure, no client id

Every SharePoint call goes through SPFx's `SPHttpClient` as the signed-in
user. Email goes through SPFx's built-in Microsoft Graph client: the package
requests the delegated **Mail.Send** permission, which a SharePoint or
Global administrator approves once under **SharePoint admin center > Advanced
> API access**. No app registration, secret or certificate is involved.

If Mail.Send is not approved, the web part falls back to SharePoint's
`SP.Utilities.Utility.SendEmail` (internal recipients only). If neither
works, the action itself still succeeds and the UI shows a warning; use
**Copy link** to send the request yourself.

## Get started

1. Download [`releases/esignature-webpart.sppkg`](releases/esignature-webpart.sppkg),
   or build it yourself (see *Build*).
2. Upload it to the tenant **App Catalog** and select **Deploy**. Choose
   *Make this solution available to all sites* to get the **Request
   signatures** button in every document library.
3. Approve the **Microsoft Graph / Mail.Send** request under SharePoint admin
   center > Advanced > API access.
4. On the site that will hold signature requests, create a page called
   **eSignature** (`SitePages/eSignature.aspx`) and add the **eSignature** web
   part.
5. Open the page as a site owner and select **Set up eSignature**.
6. Give everyone who sends or signs at least **Contribute** permission on
   that site.

The command bar button opens `SitePages/eSignature.aspx` on the site of the
library it is clicked in. To send everything to one central eSignature
site, change the `pageUrl` property of the extension to an absolute URL
(Tenant Wide Extensions list in the App Catalog for tenant-wide deployment,
or the custom action for a single site). The web part's own **Signing page
URL** property sets the link used in emails if the page URL differs from
where the web part is shown (for example in Teams).

## Using it

1. **Send**: *New request* (or **Request signatures** in a library) > choose
   the document > add signers and CCs, optionally set the signing order >
   place fields for each signer > set expiry, reminders and options > *Send*.
2. **Sign**: recipients get an email; the link opens the request, the
   fields waiting for them glow in their colour. Fill them, tick the consent
   box and *Finish and sign*. They can also decline or assign to someone
   else.
3. **Track**: the dashboard shows what needs you and what you are waiting
   for; each request shows its recipients, history and fingerprints.
4. **Complete**: when the last signer finishes, the signed PDF with its
   certificate is generated, saved and emailed to everyone.
5. **Verify**: drop a PDF on *Verify* to see whether it is the exact signed
   (or original) document of a request on this site.

## Limitations

Read these before relying on the solution for anything with legal weight.

- **Signature type.** This produces a *simple electronic signature* (SES)
  backed by an audit trail and SHA-256 fingerprints, the same category as
  most click-to-sign services. It does not apply a certificate-based PDF
  digital signature (PAdES), so it is not an *advanced* or *qualified*
  electronic signature under eIDAS. Check what your use case requires.
- **Signers need SharePoint access.** Every signer signs in with an account
  in your tenant and needs Contribute on the eSignature site. External people
  must be invited as guests. There is no email one-time-passcode flow.
- **Who can read requests.** Anyone who can read the eSignature site can
  read its lists, including signature images and the "private" per-recipient
  note. Host eSignature on a dedicated site whose membership matches who
  should see the requests, rather than a large communication site.
- **Site owners can bypass the rules.** People with Full Control can change
  list settings, edit items or delete them. Edited signature entries are
  detected and ignored, but a site owner deleting a request or its history
  cannot be prevented. Keep ownership of the eSignature site tight and use
  retention policies if you need records to be immutable.
- **Automatic reminders run when the sender opens the dashboard** (there is
  no server-side scheduler in SPFx). For unattended reminders, build a small
  Power Automate flow on the Activity list.
- **Word conversion** uses SharePoint's own PDF conversion via the v2.0
  drive API. If it is unavailable in your tenant, save the document as PDF
  first.
- **Stamped text** uses the standard Helvetica font, which covers Western
  European characters only; others appear as `?` in typed text fields (drawn,
  typed and uploaded signatures are images and are not affected).
- **Rotated pages.** Fields on PDF pages that carry a page rotation may be
  stamped in the wrong place. Re-save such documents without rotation.
- **Simultaneous last signers.** If two signers in the final parallel step
  finish at the same moment, two signed copies can be generated; the first
  one recorded is the official one.
- **List size.** The dashboard reads the 2,000 most recent requests on the
  site.

## Project layout

```
src/
  extensions/requestSignature/
    RequestSignatureCommandSet.ts     "Request signatures" library button
  webparts/eSignature/
    ESignatureWebPart.ts              Entry point and property pane
    models/types.ts                   Request, recipient, field, activity shapes
    services/
      status.ts                       Derives request state from activities (all signing rules)
      ESignService.ts                 Send, sign, decline, delegate, remind, cancel, finalise, verify
      PdfService.ts                   pdf.js loading, signed PDF + certificate of completion
      DocumentService.ts              File browse/read/upload, Word to PDF conversion
      NotificationService.ts          Email via Graph with SharePoint fallback
      Provisioning.ts                 Creates the lists and library, sets item-level edit rights
      SpRest.ts                       SPHttpClient wrapper with throttling retries
      hash.ts                         SHA-256 and ids
    components/
      ESignApp.tsx                    Navigation and deep links (?esignRequest=, ?esignSource=)
      Dashboard.tsx                   KPIs, filters, search, request list
      SendWizard.tsx                  Document > recipients > fields > review
      DocumentPicker.tsx              Library browser and upload
      FieldDesigner.tsx               Drag-and-drop field placement
      SigningView.tsx                 Signing experience, decline, delegate
      SignatureDialog.tsx             Draw / type / upload a signature
      RequestView.tsx                 Details, history, integrity, sender actions
      Templates.tsx, Verify.tsx, Setup.tsx
      PdfPages.tsx, FieldOverlay.tsx, PersonInput.tsx, shared.tsx
sharepoint/assets/                    Command set registration (site and tenant-wide)
```

## Build

Node.js 18 (SPFx 1.18 does not support Node 19+).

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

The package is written to `sharepoint/solution/esignature-webpart.sppkg`.

To try it in the hosted workbench (`https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`):

```bash
gulp serve
```
