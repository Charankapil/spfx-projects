import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  TextField
} from '@fluentui/react';

import { IField } from '../models/types';
import { CONSENT_TEXT, ESignService, IRequestWithStatus } from '../services/ESignService';
import { ISlotState } from '../services/status';
import { fieldBox, FieldValue, signedValuesOverlay } from './FieldOverlay';
import { IPageBox, PdfPages } from './PdfPages';
import { PersonInput } from './PersonInput';
import { rememberSignature, savedSignature, SignatureDialog } from './SignatureDialog';
import { errorText, FIELD_META, formatDate, recipientColour, tint } from './shared';
import styles from './ESign.module.scss';

export interface ISigningViewProps {
  service: ESignService;
  item: IRequestWithStatus;
  slot: ISlotState;
  bytes: ArrayBuffer;
  onDone: (message: string, warnings: string[]) => void;
}

const today = (): string => new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' });

export function SigningView(props: ISigningViewProps): JSX.Element {
  const { service, item, slot } = props;
  const request = item.request;
  const me = service.me;
  const myFields = React.useMemo(
    () =>
      request.payload.fields
        .filter(f => f.recipientId === slot.original.id)
        .sort((a, b) => a.page - b.page || a.y - b.y || a.x - b.x),
    [request, slot]
  );
  const [values, setValues] = React.useState<{ [id: string]: string }>(() => {
    const initial: { [id: string]: string } = {};
    myFields.forEach(f => {
      if (f.type === 'date') {
        initial[f.id] = today();
      } else if (f.type === 'name') {
        initial[f.id] = slot.current.name || me.name;
      } else if (f.type === 'email') {
        initial[f.id] = me.email;
      }
    });
    return initial;
  });
  const [adopted, setAdopted] = React.useState<{ signature?: string; initials?: string }>(() => ({
    signature: savedSignature('signature', me.email),
    initials: savedSignature('initials', me.email)
  }));
  const [padFor, setPadFor] = React.useState<IField>();
  const [focus, setFocus] = React.useState<IField>();
  const [consent, setConsent] = React.useState(false);
  const [busy, setBusy] = React.useState<string>();
  const [error, setError] = React.useState<string>();
  const [dialog, setDialog] = React.useState<'decline' | 'delegate'>();
  const [reason, setReason] = React.useState('');
  const [delegateTo, setDelegateTo] = React.useState({ name: '', email: '' });

  React.useEffect(() => {
    service.markViewed(item, slot).catch(() => undefined);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const set = (id: string, value: string | undefined): void => setValues(v => ({ ...v, [id]: value || '' }));
  const remaining = myFields.filter(f => f.required && !values[f.id]);
  const colour = recipientColour(item.status.slots.indexOf(slot));

  const activate = (field: IField): void => {
    setFocus(field);
    if (field.type === 'signature' || field.type === 'initials') {
      const existing = adopted[field.type];
      if (existing && !values[field.id]) {
        set(field.id, existing);
      } else {
        setPadFor(field);
      }
    } else if (field.type === 'checkbox') {
      set(field.id, values[field.id] === 'true' ? '' : 'true');
    }
  };

  const next = (): void => {
    const target = remaining[0] || myFields.find(f => !values[f.id]);
    if (target) {
      setFocus(target);
    }
  };

  const finish = async (): Promise<void> => {
    setError(undefined);
    setBusy('Signing...');
    try {
      // Date fields carry the moment of signing, not the moment the page was opened.
      const final = { ...values };
      myFields.filter(f => f.type === 'date').forEach(f => (final[f.id] = today()));
      const result = await service.sign(item, slot, final);
      props.onDone(
        result.completed
          ? 'You signed the document. Everyone has now signed, and the completed document has been saved and sent to all parties.'
          : 'You signed the document. The other parties have been notified.',
        result.warnings
      );
    } catch (e) {
      setError(errorText(e));
      setBusy(undefined);
    }
  };

  const decline = async (): Promise<void> => {
    setBusy('Declining...');
    try {
      const result = await service.decline(item, slot, reason.trim());
      props.onDone('You declined to sign. The sender has been notified.', result.warnings);
    } catch (e) {
      setError(errorText(e));
      setBusy(undefined);
    }
    setDialog(undefined);
  };

  const delegate = async (): Promise<void> => {
    setBusy('Delegating...');
    try {
      const result = await service.delegate(item, slot, { toName: delegateTo.name, toEmail: delegateTo.email.trim(), reason: reason.trim() });
      props.onDone(`${delegateTo.name || delegateTo.email} has been asked to sign in your place.`, result.warnings);
    } catch (e) {
      setError(errorText(e));
      setBusy(undefined);
    }
    setDialog(undefined);
  };

  const overlay = (page: IPageBox): React.ReactNode => (
    <>
      {signedValuesOverlay(request, item.status, page, slot.original.id)}
      {myFields
        .filter(f => f.page === page.index)
        .map(f => {
          const value = values[f.id];
          const editable = f.type === 'text';
          return (
            <div
              key={f.id}
              role={editable ? undefined : 'button'}
              tabIndex={editable ? undefined : 0}
              aria-label={`${FIELD_META[f.type].label}${f.required ? ' (required)' : ''}`}
              className={`${styles.field} ${value ? styles.fieldDone : styles.fieldTodo} ${focus?.id === f.id ? styles.fieldFocus : ''}`}
              style={{ ...fieldBox(f, page), borderColor: colour, background: tint(colour, 0.12), color: colour }}
              onClick={() => !editable && activate(f)}
              onKeyDown={e => !editable && (e.key === 'Enter' || e.key === ' ') && activate(f)}
              title={f.label || FIELD_META[f.type].label}
            >
              {editable ? (
                <input
                  className={styles.fieldInput}
                  value={value || ''}
                  placeholder={f.label || 'Type here'}
                  onChange={e => set(f.id, e.target.value)}
                  onFocus={() => setFocus(f)}
                  aria-label={f.label || 'Text'}
                />
              ) : value ? (
                <FieldValue field={f} value={value} page={page} />
              ) : (
                <span>
                  {f.type === 'checkbox' ? '' : `${FIELD_META[f.type].label}${f.required ? ' *' : ''}`}
                </span>
              )}
            </div>
          );
        })}
    </>
  );

  return (
    <div>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      <div className={styles.split}>
        <PdfPages bytes={props.bytes} renderOverlay={overlay} scrollToPage={focus?.page} />
        <div className={`${styles.card} ${styles.side}`}>
          <h3 className={styles.cardTitle}>{request.payload.title}</h3>
          <p className={styles.muted}>
            From {request.payload.senderName} - sent {formatDate(request.payload.sentOn)}
            {request.payload.expiresOn ? ` - sign by ${formatDate(request.payload.expiresOn)}` : ''}
          </p>
          {request.payload.message && <p style={{ whiteSpace: 'pre-wrap' }}>{request.payload.message}</p>}
          {slot.current.privateMessage && <MessageBar>{slot.current.privateMessage}</MessageBar>}
          <p>
            <strong>{myFields.length - remaining.length}</strong> of {myFields.length} fields complete
            {remaining.length > 0 ? ` - ${remaining.length} required left` : ''}
          </p>
          <div className={styles.progress}>
            <div className={styles.progressBar} style={{ width: `${myFields.length ? ((myFields.length - remaining.length) / myFields.length) * 100 : 100}%` }} />
          </div>
          <div className={styles.toolbar} style={{ marginTop: 12 }}>
            <DefaultButton iconProps={{ iconName: 'Down' }} text={remaining.length ? 'Next required field' : 'Next field'} onClick={next} />
            {(adopted.signature || adopted.initials) && (
              <DefaultButton
                text="Change my signature"
                onClick={() => {
                  const field = myFields.find(f => f.type === 'signature') || myFields.find(f => f.type === 'initials');
                  if (field) {
                    setPadFor(field);
                  }
                }}
              />
            )}
          </div>
          <Checkbox styles={{ root: { marginTop: 16 } }} label={CONSENT_TEXT} checked={consent} onChange={(_, v) => setConsent(!!v)} />
          {busy ? (
            <Spinner label={busy} styles={{ root: { marginTop: 16 } }} />
          ) : (
            <>
              <PrimaryButton
                styles={{ root: { marginTop: 16, width: '100%' } }}
                iconProps={{ iconName: 'InsertSignatureLine' }}
                text="Finish and sign"
                disabled={remaining.length > 0 || !consent}
                onClick={finish}
              />
              <div className={styles.toolbar} style={{ marginTop: 8 }}>
                <DefaultButton text="Decline" iconProps={{ iconName: 'Blocked' }} onClick={() => setDialog('decline')} />
                <DefaultButton text="Assign to someone else" iconProps={{ iconName: 'Forward' }} onClick={() => setDialog('delegate')} />
              </div>
            </>
          )}
          <p className={styles.muted} style={{ marginTop: 16 }}>
            Your signature is recorded against your SharePoint account ({me.email}) together with the fingerprint of this document.
          </p>
        </div>
      </div>

      {padFor && (
        <SignatureDialog
          kind={padFor.type === 'initials' ? 'initials' : 'signature'}
          defaultText={
            padFor.type === 'initials'
              ? (slot.current.name || me.name)
                  .split(/\s+/)
                  .map(p => p.charAt(0).toUpperCase())
                  .join('')
              : slot.current.name || me.name
          }
          onDismiss={() => setPadFor(undefined)}
          onAdopt={(dataUrl, remember) => {
            const kind = padFor.type === 'initials' ? 'initials' : 'signature';
            setAdopted(a => ({ ...a, [kind]: dataUrl }));
            rememberSignature(kind, me.email, remember ? dataUrl : undefined);
            // Apply to every field of this kind - one adoption signs them all.
            setValues(v => {
              const updated = { ...v };
              myFields.filter(f => f.type === padFor.type).forEach(f => (updated[f.id] = dataUrl));
              return updated;
            });
            setPadFor(undefined);
          }}
        />
      )}

      {dialog && (
        <Dialog
          hidden={false}
          onDismiss={() => setDialog(undefined)}
          dialogContentProps={{
            type: DialogType.normal,
            title: dialog === 'decline' ? 'Decline to sign' : 'Assign to someone else',
            subText:
              dialog === 'decline'
                ? 'Declining closes the request for everyone. The sender will be told your reason.'
                : 'The person you choose will be asked to sign instead of you. You will no longer be able to sign this request.'
          }}
          minWidth={Math.min(480, window.innerWidth - 32)}
        >
          {dialog === 'delegate' && (
            <>
              <PersonInput service={service} label="Name" required name={delegateTo.name} email={delegateTo.email} onChange={setDelegateTo} />
              <TextField label="Email" required value={delegateTo.email} onChange={(_, v) => setDelegateTo(d => ({ ...d, email: v || '' }))} />
            </>
          )}
          <TextField label={dialog === 'decline' ? 'Reason' : 'Note (optional)'} multiline rows={3} required={dialog === 'decline'} value={reason} onChange={(_, v) => setReason(v || '')} />
          <DialogFooter>
            <PrimaryButton
              text={dialog === 'decline' ? 'Decline' : 'Assign'}
              disabled={!!busy || (dialog === 'decline' ? !reason.trim() : !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(delegateTo.email.trim()))}
              onClick={dialog === 'decline' ? decline : delegate}
            />
            <DefaultButton text="Cancel" onClick={() => setDialog(undefined)} />
          </DialogFooter>
        </Dialog>
      )}
    </div>
  );
}
