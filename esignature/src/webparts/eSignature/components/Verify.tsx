import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, Spinner } from '@fluentui/react';

import { ICompletedPayload } from '../models/types';
import { ESignService, IRequestWithStatus } from '../services/ESignService';
import { sha256Hex } from '../services/hash';
import { errorText, formatDate, StatusBadge } from './shared';
import styles from './ESign.module.scss';

export interface IVerifyProps {
  service: ESignService;
  onOpen: (requestId: string) => void;
}

/**
 * Check a PDF someone sends you: its SHA-256 is computed in the browser (the
 * file is never uploaded) and looked up in the activity log. A match proves
 * the file is byte-for-byte the document that was sent or signed here.
 */
export function Verify(props: IVerifyProps): JSX.Element {
  const [busy, setBusy] = React.useState(false);
  const [result, setResult] = React.useState<{ name: string; hash: string; matches: IRequestWithStatus[] }>();
  const [error, setError] = React.useState<string>();

  const check = async (file: File | undefined): Promise<void> => {
    if (!file) {
      return;
    }
    setBusy(true);
    setError(undefined);
    setResult(undefined);
    try {
      const hash = await sha256Hex(await file.arrayBuffer());
      const matches = await props.service.findByHash(hash);
      setResult({ name: file.name, hash, matches });
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.card}>
      <h3 className={styles.cardTitle}>Verify a signed document</h3>
      <p className={styles.muted}>
        Check whether a PDF is exactly the document that was signed here. The file stays on your computer: only its fingerprint (SHA-256) is
        compared with the audit trail.
      </p>
      <label className={styles.dropzone} style={{ display: 'block' }}>
        <Icon iconName="Shield" style={{ fontSize: 28 }} />
        <div>Choose a PDF to verify</div>
        <input type="file" accept=".pdf" style={{ display: 'none' }} onChange={e => check(e.target.files?.[0])} />
      </label>
      {busy && <Spinner label="Checking..." styles={{ root: { marginTop: 16 } }} />}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {result && (
        <div style={{ marginTop: 16 }}>
          <div className={styles.muted}>{result.name}</div>
          <div className={styles.hash}>{result.hash}</div>
          {result.matches.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { marginTop: 12 } }}>
              No match. This file was not sent or signed through eSignature on this site, or it has been changed since (even re-saving a PDF
              changes its fingerprint).
            </MessageBar>
          ) : (
            result.matches.map(m => {
              const completedActivity = m.status.completed;
              const completed = completedActivity?.payload as ICompletedPayload | undefined;
              const isSigned = !!completedActivity && completed?.finalHash === result.hash;
              return (
                <MessageBar
                  key={m.request.requestId}
                  messageBarType={isSigned ? MessageBarType.success : MessageBarType.info}
                  styles={{ root: { marginTop: 12 } }}
                  actions={<DefaultButton text="Open request" onClick={() => props.onOpen(m.request.requestId)} />}
                >
                  {isSigned
                    ? `Authentic: this is the signed copy of "${m.request.payload.title}", completed ${formatDate(completedActivity?.created, true)} by ${m.status.slots
                        .map(s => s.current.name)
                        .join(', ')}.`
                    : `This is the original (unsigned) document of "${m.request.payload.title}", sent by ${m.request.payload.senderName} on ${formatDate(
                        m.request.payload.sentOn
                      )}.`}{' '}
                  <StatusBadge state={m.status.state} />
                </MessageBar>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}
