import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  MessageBar,
  MessageBarType,
  Persona,
  PersonaSize,
  PrimaryButton,
  Spinner,
  TextField
} from '@fluentui/react';

import { ICompletedPayload } from '../models/types';
import { fileName, folderOf } from '../services/DocumentService';
import { ESignService, IRequestWithStatus } from '../services/ESignService';
import { EVENT_LABELS } from '../services/PdfService';
import { sameEmail, slotForUser } from '../services/status';
import { signedValuesOverlay } from './FieldOverlay';
import { PdfPages } from './PdfPages';
import { SigningView } from './SigningView';
import { downloadBlob, errorText, formatDate, recipientColour, StatusBadge } from './shared';
import styles from './ESign.module.scss';

export interface IRequestViewProps {
  service: ESignService;
  requestId: string;
  onBack: () => void;
}

interface INotice {
  type: MessageBarType;
  text: string;
}

function csvCell(value: string): string {
  return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

export function RequestView(props: IRequestViewProps): JSX.Element {
  const { service, requestId } = props;
  const [item, setItem] = React.useState<IRequestWithStatus | null>();
  const [bytes, setBytes] = React.useState<ArrayBuffer>();
  const [signing, setSigning] = React.useState(true);
  const [busy, setBusy] = React.useState<string>();
  const [notices, setNotices] = React.useState<INotice[]>([]);
  const [cancelOpen, setCancelOpen] = React.useState(false);
  const [reason, setReason] = React.useState('');

  const load = React.useCallback(async () => {
    try {
      const loaded = await service.loadOne(requestId);
      setItem(loaded || null);
      if (loaded) {
        const done = loaded.status.completed?.payload as ICompletedPayload | undefined;
        setBytes(await service.docs.readLocal(done ? done.signedUrl : loaded.request.payload.documentUrl));
      }
    } catch (e) {
      setNotices([{ type: MessageBarType.error, text: errorText(e) }]);
      setItem(null);
    }
  }, [service, requestId]);

  React.useEffect(() => {
    load().catch(() => undefined);
  }, [load]);

  const run = async (label: string, action: () => Promise<{ warnings: string[] } | void>, success: string): Promise<void> => {
    setBusy(label);
    try {
      const result = await action();
      const warnings = (result && result.warnings) || [];
      setNotices([{ type: MessageBarType.success, text: success }].concat(warnings.map(w => ({ type: MessageBarType.warning, text: w }))));
      await load();
    } catch (e) {
      setNotices([{ type: MessageBarType.error, text: errorText(e) }]);
    } finally {
      setBusy(undefined);
    }
  };

  if (item === undefined) {
    return <Spinner label="Loading request..." styles={{ root: { padding: 32 } }} />;
  }
  if (item === null) {
    return (
      <div className={styles.card}>
        {notices.map((n, i) => (
          <MessageBar key={i} messageBarType={n.type}>
            {n.text}
          </MessageBar>
        ))}
        <p>This request could not be found, or you do not have access to it.</p>
        <DefaultButton text="Back" onClick={props.onBack} />
      </div>
    );
  }

  const { request, status, activities } = item;
  const p = request.payload;
  const me = service.me;
  const isSender = sameEmail(request.authorEmail, me.email);
  const mySlot = slotForUser(status, me.email);
  const completed = status.completed?.payload as ICompletedPayload | undefined;
  const copied = activities.some(a => a.type === 'copied');

  const noticeBars = notices.map((n, i) => (
    <MessageBar key={i} messageBarType={n.type} onDismiss={() => setNotices(list => list.filter(x => x !== n))}>
      {n.text}
    </MessageBar>
  ));

  if (mySlot && signing && bytes) {
    return (
      <div>
        <div className={styles.toolbar}>
          <DefaultButton iconProps={{ iconName: 'Back' }} text="Back" onClick={props.onBack} />
          <DefaultButton iconProps={{ iconName: 'History' }} text="Details and history" onClick={() => setSigning(false)} />
        </div>
        {noticeBars}
        <SigningView
          service={service}
          item={item}
          slot={mySlot}
          bytes={bytes}
          onDone={(message, warnings) => {
            setSigning(false);
            setNotices([{ type: MessageBarType.success, text: message }].concat(warnings.map(w => ({ type: MessageBarType.warning, text: w }))));
            setItem(undefined);
            load().catch(() => undefined);
          }}
        />
      </div>
    );
  }

  const exportAudit = (): void => {
    const rows = [['Time (UTC)', 'Event', 'By', 'Email', 'For recipient', 'Document SHA-256', 'Details']];
    activities
      .slice()
      .sort((a, b) => new Date(a.created).getTime() - new Date(b.created).getTime())
      .forEach(a => {
        const slot = status.slots.find(s => s.original.id === a.slot);
        const payload = a.payload as { reason?: string; toEmail?: string; automatic?: boolean };
        const details = payload?.reason ? `Reason: ${payload.reason}` : payload?.toEmail ? `To: ${payload.toEmail}` : payload?.automatic ? 'Automatic' : '';
        rows.push([a.created, EVENT_LABELS[a.type] || a.type, a.authorName, a.authorEmail, slot ? slot.current.email : '', a.docHash || '', details]);
      });
    downloadBlob(rows.map(r => r.map(csvCell).join(',')).join('\r\n'), `audit-${requestId}.csv`, 'text/csv');
  };

  const openUrl = (url: string): void => {
    window.open(`${new URL(service.rest.webUrl).origin}${encodeURI(url)}`, '_blank', 'noopener');
  };

  return (
    <div>
      <div className={styles.toolbar}>
        <DefaultButton iconProps={{ iconName: 'Back' }} text="Back" onClick={props.onBack} />
        {mySlot && <PrimaryButton iconProps={{ iconName: 'InsertSignatureLine' }} text="Review and sign" onClick={() => setSigning(true)} />}
        <span className={styles.spacer} />
        {completed && <PrimaryButton iconProps={{ iconName: 'Download' }} text="Signed document" onClick={() => openUrl(completed.signedUrl)} />}
        <DefaultButton iconProps={{ iconName: 'PDF' }} text="Original" onClick={() => openUrl(p.documentUrl)} />
        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Audit (CSV)" onClick={exportAudit} />
        <DefaultButton
          iconProps={{ iconName: 'Link' }}
          text="Copy link"
          onClick={() =>
            navigator.clipboard
              .writeText(service.link(requestId))
              .then(() => setNotices([{ type: MessageBarType.info, text: 'Link copied. Only the recipients can sign with it.' }]))
              .catch(() => setNotices([{ type: MessageBarType.info, text: service.link(requestId) }]))
          }
        />
      </div>
      {noticeBars}
      {busy && <Spinner label={busy} styles={{ root: { margin: 12 } }} />}

      <div className={styles.split}>
        <div>
          <div className={styles.card}>
            <div className={styles.toolbar}>
              <h3 className={styles.cardTitle} style={{ margin: 0 }}>
                {p.title}
              </h3>
              <StatusBadge state={status.state} actionRequired={!!mySlot} />
            </div>
            <p className={styles.muted}>
              {p.documentName} - sent by {p.senderName} on {formatDate(p.sentOn, true)}
              {p.expiresOn ? ` - expires ${formatDate(p.expiresOn)}` : ''}
              {p.reminderDays ? ` - reminders every ${p.reminderDays} day(s)` : ''}
            </p>
            {p.message && <p style={{ whiteSpace: 'pre-wrap' }}>{p.message}</p>}
            {status.declined && (
              <MessageBar messageBarType={MessageBarType.error}>
                Declined by {status.declined.authorName}: {(status.declined.payload as { reason?: string }).reason || 'no reason given'}
              </MessageBar>
            )}
            {status.cancelled && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Cancelled by the sender {formatDate(status.cancelled.created, true)}
                {(status.cancelled.payload as { reason?: string }).reason ? `: ${(status.cancelled.payload as { reason?: string }).reason}` : ''}
              </MessageBar>
            )}
            {isSender && status.state === 'InProgress' && (
              <div className={styles.toolbar} style={{ marginTop: 12 }}>
                <DefaultButton
                  iconProps={{ iconName: 'Ringer' }}
                  text="Send reminder now"
                  disabled={!!busy}
                  onClick={() => run('Sending reminders...', () => service.remind(item), 'Reminder sent.')}
                />
                <DefaultButton iconProps={{ iconName: 'Cancel' }} text="Cancel request" disabled={!!busy} onClick={() => setCancelOpen(true)} />
              </div>
            )}
            {status.state === 'Completed' && !status.completed && (
              <MessageBar
                messageBarType={MessageBarType.warning}
                actions={
                  <PrimaryButton
                    text="Generate signed document"
                    disabled={!!busy}
                    onClick={() => run('Generating the signed document...', () => service.finalize(item), 'Signed document generated and sent to everyone.')}
                  />
                }
              >
                Everyone has signed, but the signed document has not been generated yet.
              </MessageBar>
            )}
            {isSender && completed && p.sourceUrl && !copied && (
              <MessageBar
                actions={
                  <DefaultButton
                    text="Save copy next to original"
                    disabled={!!busy}
                    onClick={() => run('Saving copy...', () => service.copyToSource(request), 'Signed copy saved next to the original document.')}
                  />
                }
              >
                The original is in {folderOf(p.sourceUrl)}.
              </MessageBar>
            )}
          </div>

          <div className={styles.card}>
            <h3 className={styles.cardTitle}>
              Recipients - {status.signedCount} of {status.signerCount} signed
            </h3>
            <ul className={styles.signerList}>
              {status.slots.map((s, i) => (
                <li key={s.original.id} className={styles.signerItem}>
                  <span className={styles.recipientDot} style={{ background: recipientColour(i) }} />
                  <Persona
                    text={s.current.name}
                    secondaryText={
                      s.status === 'signed'
                        ? `Signed ${formatDate(s.signedOn, true)}`
                        : s.status === 'declined'
                        ? 'Declined'
                        : s.status === 'pending'
                        ? s.viewedOn
                          ? `Viewed ${formatDate(s.viewedOn, true)} - needs to sign`
                          : 'Needs to sign'
                        : `Waiting for step ${s.original.order}`
                    }
                    tertiaryText={s.delegations.length ? `Delegated by ${s.original.name}` : s.current.email}
                    size={PersonaSize.size40}
                    styles={{ root: { flex: 1 } }}
                  />
                  <Icon
                    iconName={s.status === 'signed' ? 'CompletedSolid' : s.status === 'declined' ? 'StatusErrorFull' : s.status === 'pending' ? 'Clock' : 'More'}
                    style={{ color: s.status === 'signed' ? '#107c10' : s.status === 'declined' ? '#a4262c' : '#8a8886', fontSize: 18 }}
                  />
                </li>
              ))}
              {status.ccs.map(c => (
                <li key={c.id} className={styles.signerItem}>
                  <span className={styles.recipientDot} style={{ background: '#8a8886' }} />
                  <Persona text={c.name} secondaryText="Receives a copy" tertiaryText={c.email} size={PersonaSize.size40} />
                </li>
              ))}
            </ul>
          </div>

          {bytes && (
            <PdfPages bytes={bytes} renderOverlay={page => (completed ? null : signedValuesOverlay(request, status, page))} />
          )}
        </div>

        <div className={styles.side}>
          <div className={styles.card}>
            <h3 className={styles.cardTitle}>History</h3>
            <ul className={styles.timeline}>
              {activities
                .slice()
                .sort((a, b) => new Date(a.created).getTime() - new Date(b.created).getTime())
                .map(a => {
                  const payload = a.payload as { reason?: string; toName?: string; toEmail?: string; automatic?: boolean };
                  return (
                    <li key={a.id} className={styles.timelineItem}>
                      <strong>{EVENT_LABELS[a.type] || a.type}</strong>
                      {a.type === 'delegated' && payload.toEmail ? ` to ${payload.toName || payload.toEmail}` : ''}
                      {a.type === 'reminded' && payload.automatic ? ' (automatic)' : ''}
                      <br />
                      <span className={styles.muted}>
                        {a.authorName} - {formatDate(a.created, true)}
                      </span>
                      {status.ignored[a.id] && <div style={{ color: '#a4262c', fontSize: 12 }}>Not counted: {status.ignored[a.id]}</div>}
                      {payload.reason && <div className={styles.muted}>&ldquo;{payload.reason}&rdquo;</div>}
                    </li>
                  );
                })}
            </ul>
          </div>
          <div className={styles.card}>
            <h3 className={styles.cardTitle}>Integrity</h3>
            <div className={styles.muted}>Original document SHA-256</div>
            <div className={styles.hash}>{p.documentHash}</div>
            {completed && (
              <>
                <div className={styles.muted} style={{ marginTop: 8 }}>
                  Signed document SHA-256 ({fileName(completed.signedUrl)})
                </div>
                <div className={styles.hash}>{completed.finalHash}</div>
              </>
            )}
            {status.slots
              .map(s => {
                if (!s.signedActivity) {
                  return null;
                }
                const hash = (s.signedActivity.payload as { documentHash?: string }).documentHash;
                return (
                  <div key={s.original.id} className={styles.muted} style={{ marginTop: 8 }}>
                    <Icon iconName={hash === p.documentHash ? 'CheckMark' : 'Warning'} /> {s.current.name}{' '}
                    {hash === p.documentHash ? 'signed this exact document' : 'signed a document with a different fingerprint'}
                  </div>
                );
              })}
            {status.warnings.map((w, i) => (
              <MessageBar key={i} messageBarType={MessageBarType.severeWarning} styles={{ root: { marginTop: 8 } }}>
                {w}
              </MessageBar>
            ))}
          </div>
        </div>
      </div>

      {cancelOpen && (
        <Dialog
          hidden={false}
          onDismiss={() => setCancelOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Cancel this request?',
            subText: 'Nobody will be able to sign it any more. Recipients who still had to sign are told it was cancelled.'
          }}
        >
          <TextField label="Reason (optional)" multiline rows={3} value={reason} onChange={(_, v) => setReason(v || '')} />
          <DialogFooter>
            <PrimaryButton
              text="Cancel request"
              onClick={() => {
                setCancelOpen(false);
                run('Cancelling...', () => service.cancel(item, reason.trim()), 'The request has been cancelled.').catch(() => undefined);
              }}
            />
            <DefaultButton text="Keep it" onClick={() => setCancelOpen(false)} />
          </DialogFooter>
        </Dialog>
      )}
    </div>
  );
}
