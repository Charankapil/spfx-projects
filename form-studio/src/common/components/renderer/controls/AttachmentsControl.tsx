import * as React from 'react';
import { ActionButton, IconButton } from '@fluentui/react';

import { IAttachmentsValue } from '../../../form/formValues';
import styles from '../FormRenderer.module.scss';
import { IFieldControlProps } from './types';

export const AttachmentsControl: React.FC<IFieldControlProps> = (props) => {
  const value = (props.value as IAttachmentsValue) || { existing: [], removed: [], added: [] };
  const input = React.useRef<HTMLInputElement>(null);
  const readOnly = props.state.readOnly;

  const kept = value.existing.filter((a) => value.removed.indexOf(a.fileName) < 0);

  return (
    <div className={styles.attachments}>
      {kept.map((a) => (
        <div key={a.fileName} className={styles.attachment}>
          <a href={a.serverRelativeUrl} target="_blank" rel="noopener noreferrer" data-interception="off">
            {a.fileName}
          </a>
          {!readOnly && (
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title={`Remove ${a.fileName}`}
              ariaLabel={`Remove ${a.fileName}`}
              onClick={() => props.onChange({ ...value, removed: value.removed.concat([a.fileName]) })}
            />
          )}
        </div>
      ))}
      {value.added.map((f, i) => (
        <div key={`new-${i}`} className={styles.attachment}>
          <span>{f.name}</span>
          <span className={styles.muted}>(new)</span>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            title={`Remove ${f.name}`}
            ariaLabel={`Remove ${f.name}`}
            onClick={() => props.onChange({ ...value, added: value.added.filter((_x, j) => j !== i) })}
          />
        </div>
      ))}
      {readOnly && !kept.length && <span className={styles.muted}>No attachments</span>}
      {!readOnly && (
        <>
          <input
            ref={input}
            id={props.id}
            type="file"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              const files: File[] = [];
              const list = e.target.files;
              for (let i = 0; list && i < list.length; i++) {
                files.push(list[i]);
              }
              props.onChange({ ...value, added: value.added.concat(files) });
              e.target.value = '';
            }}
          />
          <ActionButton iconProps={{ iconName: 'Attach' }} onClick={() => input.current && input.current.click()}>
            Attach files
          </ActionButton>
        </>
      )}
    </div>
  );
};
