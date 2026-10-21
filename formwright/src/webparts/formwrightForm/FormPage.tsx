import * as React from 'react';
import { DefaultButton, Icon, MessageBar, PrimaryButton } from '@fluentui/react';

import { FormHost } from '../../common/components/renderer/FormHost';
import styles from '../../common/components/renderer/FormRenderer.module.scss';
import { FormMode } from '../../common/model/IFormDefinition';
import { ISpContext } from '../../common/services/http';

export interface IFormPageProps {
  ctx: ISpContext;
  listId: string;
  formId?: string;
  itemId?: number;
  mode: FormMode;
  allowAnother: boolean;
  onConfigure: () => void;
}

/** A form on a page: submit, then a confirmation (or a redirect set in the designer). */
export const FormPage: React.FC<IFormPageProps> = (props) => {
  const [submitted, setSubmitted] = React.useState<string | undefined>();
  const [mode, setMode] = React.useState<FormMode>(props.mode);
  // Bumping the key remounts the form with fresh values for "Submit another response".
  const [run, setRun] = React.useState<number>(0);

  React.useEffect(() => setMode(props.mode), [props.mode, props.itemId]);

  if (!props.listId) {
    return (
      <MessageBar actions={<PrimaryButton text="Choose a list" onClick={props.onConfigure} />}>
        Choose the list this form saves to.
      </MessageBar>
    );
  }

  if (submitted) {
    return (
      <div className={styles.success} role="status">
        <Icon iconName="CompletedSolid" className={styles.successIcon} />
        <h2>Saved</h2>
        <p>{submitted}</p>
        {props.allowAnother && props.mode === 'New' && (
          <DefaultButton
            text="Submit another response"
            onClick={() => {
              setSubmitted(undefined);
              setRun(run + 1);
            }}
          />
        )}
      </div>
    );
  }

  return (
    <FormHost
      key={`${run}-${mode}`}
      ctx={props.ctx}
      listId={props.listId}
      formId={props.formId}
      mode={mode}
      itemId={mode === 'New' ? undefined : props.itemId}
      onSaved={(_id, _values, redirectUrl, successMessage) => {
        if (redirectUrl) {
          window.location.href = redirectUrl;
        } else if (props.mode === 'Display') {
          // Opened read-only and switched to edit: go back to the read-only view of the saved item.
          setMode('Display');
          setRun(run + 1);
        } else {
          setSubmitted(successMessage || 'Your response has been saved.');
        }
      }}
      onEdit={mode === 'Display' ? () => setMode('Edit') : undefined}
      onCancel={mode === 'Edit' && props.mode === 'Display' ? () => setMode('Display') : undefined}
    />
  );
};
