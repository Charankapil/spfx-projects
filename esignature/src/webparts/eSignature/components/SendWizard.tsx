import * as React from 'react';
import {
  Checkbox,
  DatePicker,
  DefaultButton,
  Dropdown,
  IconButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SpinButton,
  Spinner,
  TextField,
  Toggle
} from '@fluentui/react';

import { IField, IRecipient, ITemplate, RecipientRole } from '../models/types';
import { baseName } from '../services/DocumentService';
import { ESignService } from '../services/ESignService';
import { shortId } from '../services/hash';
import { loadPdfLib } from '../services/PdfService';
import { sameEmail, toIsoDate } from '../services/status';
import { DocumentPicker, IPickedDocument } from './DocumentPicker';
import { FieldDesigner } from './FieldDesigner';
import { PersonInput } from './PersonInput';
import { errorText, FIELD_META, formatDate, recipientColour } from './shared';
import styles from './ESign.module.scss';

export interface ISendWizardProps {
  service: ESignService;
  sourceUrl?: string;
  template?: ITemplate;
  onSent: (requestId: string, warnings: string[]) => void;
  onCancel: () => void;
}

const STEPS = ['Document', 'Recipients', 'Fields', 'Review & send'];
const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function newRecipient(order: number, role: RecipientRole = 'signer'): IRecipient {
  return { id: shortId(), name: '', email: '', role, order };
}

export function SendWizard(props: ISendWizardProps): JSX.Element {
  const { service, template } = props;
  const [step, setStep] = React.useState(0);
  const [doc, setDoc] = React.useState<IPickedDocument>();
  const [title, setTitle] = React.useState('');
  const [message, setMessage] = React.useState(template?.message || '');
  const [recipients, setRecipients] = React.useState<IRecipient[]>(
    template ? template.recipients.map(r => ({ ...r, name: '', email: '' })) : [newRecipient(1)]
  );
  const [sequential, setSequential] = React.useState(template ? new Set(template.recipients.map(r => r.order)).size > 1 : false);
  const [fields, setFields] = React.useState<IField[]>(template?.fields || []);
  const [pageCount, setPageCount] = React.useState(0);
  const [expiresOn, setExpiresOn] = React.useState<Date | undefined>(() => {
    const d = new Date();
    d.setDate(d.getDate() + (template?.expiresInDays || 30));
    return d;
  });
  const [reminderDays, setReminderDays] = React.useState(template ? template.reminderDays : 3);
  const [saveToSource, setSaveToSource] = React.useState(true);
  const [templateName, setTemplateName] = React.useState('');
  const [saveTemplate, setSaveTemplate] = React.useState(false);
  const [progress, setProgress] = React.useState<string>();
  const [error, setError] = React.useState<string>();

  const signers = recipients.filter(r => r.role === 'signer');

  const picked = async (d: IPickedDocument): Promise<void> => {
    setError(undefined);
    try {
      const lib = await loadPdfLib();
      const pdf = await lib.PDFDocument.load(d.pdf);
      setPageCount(pdf.getPageCount());
      if (template) {
        // A template designed on a longer document: drop fields that would fall off the end.
        setFields(f => f.filter(x => x.page < pdf.getPageCount()));
      }
    } catch (e) {
      setError(/encrypt/i.test(errorText(e)) ? 'This PDF is password-protected or encrypted and cannot be signed. Remove the protection and try again.' : `This PDF could not be read: ${errorText(e)}`);
      return;
    }
    setDoc(d);
    if (!title) {
      setTitle(`Please sign: ${baseName(d.name)}`);
    }
  };

  const setRecipient = (id: string, change: Partial<IRecipient>): void =>
    setRecipients(list => list.map(r => (r.id === id ? { ...r, ...change } : r)));

  const recipientProblems = (): string | undefined => {
    if (!signers.length) {
      return 'Add at least one signer.';
    }
    for (const r of recipients) {
      if (!r.name.trim() || !EMAIL.test(r.email.trim())) {
        return `Enter a name and a valid email address for every recipient${r.roleName ? ` (${r.roleName})` : ''}.`;
      }
    }
    const dupes = signers.filter((s, i) => signers.some((o, j) => j < i && sameEmail(o.email, s.email)));
    if (dupes.length) {
      return `${dupes[0].email} is added as a signer more than once.`;
    }
    return undefined;
  };

  const fieldProblems = (): string | undefined => {
    const missing = signers.filter(s => !fields.some(f => f.recipientId === s.id && f.type === 'signature'));
    return missing.length ? `Place a signature field for ${missing.map(m => m.name || m.email).join(', ')}.` : undefined;
  };

  const problem = step === 0 ? (!doc ? 'Choose a document.' : !title.trim() ? 'Enter a title.' : undefined) : step === 1 ? recipientProblems() : step === 2 ? fieldProblems() : undefined;

  const send = async (): Promise<void> => {
    if (!doc) {
      return;
    }
    setError(undefined);
    try {
      const ordered = recipients.map(r => ({ ...r, name: r.name.trim(), email: r.email.trim(), order: sequential ? r.order : 1 }));
      // Drop fields that belong to recipients who were removed or turned into CCs.
      const validFields = fields.filter(f => ordered.some(r => r.id === f.recipientId && r.role === 'signer'));
      if (saveTemplate && templateName.trim()) {
        setProgress('Saving template...');
        await service.saveTemplate({
          name: templateName.trim(),
          description: '',
          recipients: ordered.map(r => ({ ...r, roleName: r.roleName || (r.role === 'cc' ? 'CC' : `Signer ${r.order}`), name: '', email: '' })),
          fields: validFields,
          message,
          reminderDays,
          expiresInDays: expiresOn ? Math.max(1, Math.round((expiresOn.getTime() - Date.now()) / 86400000)) : 30,
          pageCount
        });
      }
      const result = await service.send(
        {
          pdf: doc.pdf,
          payload: {
            title: title.trim(),
            message: message.trim(),
            documentName: doc.name,
            sourceUrl: doc.sourceUrl,
            pageCount,
            recipients: ordered,
            fields: validFields,
            expiresOn: expiresOn ? toIsoDate(expiresOn) : undefined,
            reminderDays,
            saveToSource: saveToSource && !!doc.sourceUrl
          }
        },
        setProgress
      );
      props.onSent(result.requestId, result.warnings);
    } catch (e) {
      setError(errorText(e));
      setProgress(undefined);
    }
  };

  return (
    <div>
      <div className={styles.steps}>
        {STEPS.map((label, i) => (
          <span key={label} className={`${styles.step} ${i === step ? styles.stepActive : ''}`}>
            <span className={styles.stepNumber}>{i + 1}</span>
            {label}
          </span>
        ))}
      </div>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}

      {step === 0 && (
        <div className={styles.card}>
          <h3 className={styles.cardTitle}>Choose a document</h3>
          {template && <MessageBar>Using template &ldquo;{template.name}&rdquo;. Choose the document to send with it.</MessageBar>}
          {doc ? (
            <MessageBar messageBarType={MessageBarType.success} actions={<DefaultButton text="Change" onClick={() => setDoc(undefined)} />}>
              {doc.name} - {pageCount} page{pageCount === 1 ? '' : 's'}
              {doc.converted ? ' (converted to PDF for signing)' : ''}
            </MessageBar>
          ) : (
            <DocumentPicker service={service} initialUrl={props.sourceUrl} onPicked={picked} />
          )}
          <TextField label="Title" required value={title} onChange={(_, v) => setTitle(v || '')} styles={{ root: { marginTop: 12 } }} />
          <TextField label="Message to recipients" multiline rows={3} value={message} onChange={(_, v) => setMessage(v || '')} />
        </div>
      )}

      {step === 1 && (
        <div className={styles.card}>
          <h3 className={styles.cardTitle}>Who needs to sign?</h3>
          <Toggle
            label="Set signing order"
            inlineLabel
            checked={sequential}
            onChange={(_, v) => setSequential(!!v)}
            onText="Signers sign in the order given (same number = sign at the same time)"
            offText="Everyone can sign at the same time"
          />
          {recipients.map((r, i) => (
            <div key={r.id} className={styles.recipientRow}>
              {sequential ? (
                <SpinButton
                  label="Order"
                  min={1}
                  max={50}
                  value={String(r.order)}
                  disabled={r.role === 'cc'}
                  onChange={(_, v) => setRecipient(r.id, { order: Math.max(1, parseInt(v || '1', 10) || 1) })}
                />
              ) : (
                <div style={{ paddingBottom: 6 }}>
                  <span className={styles.recipientDot} style={{ background: recipientColour(signers.indexOf(r)) }} />
                  {i + 1}
                </div>
              )}
              <PersonInput
                service={service}
                label={r.roleName ? `${r.roleName} - name` : 'Name'}
                required
                name={r.name}
                email={r.email}
                onChange={p => setRecipient(r.id, p)}
              />
              <TextField label="Email" required value={r.email} onChange={(_, v) => setRecipient(r.id, { email: v || '' })} />
              <Dropdown
                label="Role"
                selectedKey={r.role}
                options={[
                  { key: 'signer', text: 'Needs to sign' },
                  { key: 'cc', text: 'Receives a copy' }
                ]}
                onChange={(_, o) => setRecipient(r.id, { role: o?.key as RecipientRole })}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remove recipient"
                ariaLabel="Remove recipient"
                disabled={recipients.length === 1}
                onClick={() => {
                  setRecipients(list => list.filter(x => x.id !== r.id));
                  setFields(list => list.filter(f => f.recipientId !== r.id));
                }}
              />
            </div>
          ))}
          <div className={styles.toolbar} style={{ marginTop: 12 }}>
            <DefaultButton
              iconProps={{ iconName: 'AddFriend' }}
              text="Add signer"
              onClick={() => setRecipients(list => list.concat(newRecipient(sequential ? Math.max(0, ...list.map(x => x.order)) + 1 : 1)))}
            />
            <DefaultButton iconProps={{ iconName: 'Mail' }} text="Add CC" onClick={() => setRecipients(list => list.concat(newRecipient(1, 'cc')))} />
          </div>
        </div>
      )}

      {step === 2 && doc && <FieldDesigner bytes={doc.pdf} signers={signers} fields={fields} onChange={setFields} />}

      {step === 3 && doc && (
        <div className={styles.card}>
          <h3 className={styles.cardTitle}>Review and send</h3>
          <p>
            <strong>{title}</strong>
            <br />
            <span className={styles.muted}>
              {doc.name} - {pageCount} page{pageCount === 1 ? '' : 's'} - {fields.length} field{fields.length === 1 ? '' : 's'}
            </span>
          </p>
          <ul className={styles.signerList}>
            {recipients
              .slice()
              .sort((a, b) => (a.role === b.role ? a.order - b.order : a.role === 'signer' ? -1 : 1))
              .map(r => (
                <li key={r.id} className={styles.signerItem}>
                  <span className={styles.recipientDot} style={{ background: r.role === 'cc' ? '#8a8886' : recipientColour(signers.indexOf(r)) }} />
                  <span style={{ flex: 1 }}>
                    {r.name} &lt;{r.email}&gt;
                    <br />
                    <span className={styles.muted}>
                      {r.role === 'cc'
                        ? 'Receives a copy'
                        : `${sequential ? `Signs in step ${r.order}` : 'Signs'} - ${fields
                            .filter(f => f.recipientId === r.id)
                            .map(f => FIELD_META[f.type].label)
                            .join(', ')}`}
                    </span>
                  </span>
                </li>
              ))}
          </ul>
          <DatePicker
            label="Expires on"
            value={expiresOn}
            minDate={new Date()}
            onSelectDate={d => setExpiresOn(d || undefined)}
            formatDate={d => formatDate(d?.toISOString())}
            allowTextInput={false}
          />
          <Checkbox label="Never expires" checked={!expiresOn} onChange={(_, v) => setExpiresOn(v ? undefined : new Date(Date.now() + 30 * 86400000))} />
          <Dropdown
            label="Automatic reminders"
            selectedKey={reminderDays}
            onChange={(_, o) => setReminderDays(o?.key as number)}
            options={[
              { key: 0, text: 'No automatic reminders' },
              { key: 1, text: 'Every day' },
              { key: 2, text: 'Every 2 days' },
              { key: 3, text: 'Every 3 days' },
              { key: 7, text: 'Every week' }
            ]}
            styles={{ root: { marginTop: 12 } }}
          />
          {doc.sourceUrl && (
            <Checkbox
              styles={{ root: { marginTop: 12 } }}
              label="Save the signed copy next to the original document"
              checked={saveToSource}
              onChange={(_, v) => setSaveToSource(!!v)}
            />
          )}
          <Checkbox
            styles={{ root: { marginTop: 12 } }}
            label="Save recipients and fields as a reusable template"
            checked={saveTemplate}
            onChange={(_, v) => setSaveTemplate(!!v)}
          />
          {saveTemplate && <TextField label="Template name" required value={templateName} onChange={(_, v) => setTemplateName(v || '')} />}
          {progress && <Spinner label={progress} styles={{ root: { marginTop: 16 } }} />}
        </div>
      )}

      <div className={styles.footer}>
        {problem && step < 3 && <span className={styles.muted} style={{ alignSelf: 'center', marginRight: 'auto' }}>{problem}</span>}
        <DefaultButton text="Cancel" onClick={props.onCancel} disabled={!!progress} />
        {step > 0 && <DefaultButton text="Back" onClick={() => setStep(step - 1)} disabled={!!progress} />}
        {step < 3 && <PrimaryButton text="Next" disabled={!!problem} onClick={() => setStep(step + 1)} />}
        {step === 3 && (
          <PrimaryButton
            iconProps={{ iconName: 'Send' }}
            text="Send"
            disabled={!!progress || (saveTemplate && !templateName.trim())}
            onClick={send}
          />
        )}
      </div>
    </div>
  );
}
