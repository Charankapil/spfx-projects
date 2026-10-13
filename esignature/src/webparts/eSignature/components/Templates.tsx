import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, Spinner } from '@fluentui/react';

import { ITemplate } from '../models/types';
import { ESignService } from '../services/ESignService';
import { errorText, FIELD_META } from './shared';
import styles from './ESign.module.scss';

export interface ITemplatesProps {
  service: ESignService;
  onUse: (template: ITemplate) => void;
}

export function Templates(props: ITemplatesProps): JSX.Element {
  const [templates, setTemplates] = React.useState<ITemplate[]>();
  const [error, setError] = React.useState<string>();

  const load = React.useCallback(() => {
    props.service
      .templates()
      .then(setTemplates)
      .catch(e => setError(errorText(e)));
  }, [props.service]);

  React.useEffect(load, [load]);

  if (!templates && !error) {
    return <Spinner label="Loading templates..." styles={{ root: { padding: 32 } }} />;
  }

  return (
    <div>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {templates && templates.length === 0 && (
        <div className={`${styles.card} ${styles.empty}`}>
          No templates yet. When you send a request, tick &ldquo;Save recipients and fields as a reusable template&rdquo; on the last step.
        </div>
      )}
      {templates?.map(t => {
        const counts: { [label: string]: number } = {};
        t.fields.forEach(f => (counts[FIELD_META[f.type].label] = (counts[FIELD_META[f.type].label] || 0) + 1));
        return (
          <div key={t.itemId} className={styles.card}>
            <div className={styles.toolbar}>
              <h3 className={styles.cardTitle} style={{ margin: 0 }}>
                {t.name}
              </h3>
              <span className={styles.spacer} />
              <PrimaryButton iconProps={{ iconName: 'Send' }} text="Use" onClick={() => props.onUse(t)} />
              <DefaultButton
                iconProps={{ iconName: 'Delete' }}
                text="Delete"
                onClick={() => {
                  if (t.itemId && window.confirm(`Delete the template "${t.name}"?`)) {
                    props.service
                      .deleteTemplate(t.itemId)
                      .then(load)
                      .catch(e => setError(errorText(e)));
                  }
                }}
              />
            </div>
            <div className={styles.muted}>
              Roles:{' '}
              {t.recipients
                .map(r => `${r.roleName || (r.role === 'cc' ? 'CC' : 'Signer')}${r.role === 'signer' ? ` (step ${r.order})` : ''}`)
                .join(', ')}
            </div>
            <div className={styles.muted}>
              Fields:{' '}
              {Object.keys(counts)
                .map(k => `${counts[k]} ${k.toLowerCase()}`)
                .join(', ') || 'none'}{' '}
              - designed on a {t.pageCount}-page document
            </div>
          </div>
        );
      })}
    </div>
  );
}
