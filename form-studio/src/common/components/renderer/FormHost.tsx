import * as React from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { IRuntimeContext, selectForm } from '../../form/FormEngine';
import { IFormValues } from '../../form/formValues';
import { createDefaultForm } from '../../model/formFactory';
import { IFieldInfo } from '../../model/IFieldInfo';
import { FormMode, IFormDefinition } from '../../model/IFormDefinition';
import { FormStore } from '../../services/FormStore';
import { errorMessage, ISpContext } from '../../services/http';
import { ListService } from '../../services/ListService';
import styles from './FormRenderer.module.scss';
import { FormView } from './FormView';

export interface IFormHostProps {
  ctx: ISpContext;
  listId: string;
  mode: FormMode;
  itemId?: number;
  contentTypeId?: string;
  /** Server-relative folder to create new items in; ignored when it is the list's root. */
  folderPath?: string;
  /** Use this form instead of picking one by mode / audience / content type. */
  formId?: string;
  onSaved: (itemId: number, values: IFormValues, redirectUrl?: string, successMessage?: string) => void;
  onCancel?: () => void;
  onEdit?: () => void;
}

interface ILoaded {
  form: IFormDefinition;
  fields: IFieldInfo[];
  runtime: IRuntimeContext;
  values: IFormValues;
  etag?: string;
  contentTypeId?: string;
  folderPath?: string;
  listService: ListService;
}

function trimSlash(path: string): string {
  return path.replace(/\/+$/, '').toLowerCase();
}

/** Loads the list, the right form and the item, then renders the form. */
export const FormHost: React.FC<IFormHostProps> = (props) => {
  const { ctx, listId, mode, itemId, formId } = props;
  const [loaded, setLoaded] = React.useState<ILoaded | undefined>();
  const [error, setError] = React.useState<string | undefined>();

  React.useEffect(() => {
    let cancelled = false;
    const load = async (): Promise<void> => {
      const listService = new ListService(ctx);
      const store = new FormStore(ctx);
      const [list, fields, user] = await Promise.all([listService.getList(listId), listService.getFields(listId), listService.getCurrentUser()]);
      const saved = await store.load(listId, list.title);

      let values: IFormValues = {};
      let etag: string | undefined;
      let contentTypeId = props.contentTypeId;
      if (itemId) {
        const item = await listService.getItem(listId, itemId, fields);
        values = item.values;
        etag = item.etag;
        contentTypeId = contentTypeId || item.contentTypeId;
      }

      const form =
        (formId && saved.forms.filter((f) => f.id === formId)[0]) ||
        selectForm(saved.forms, mode, user.groups, contentTypeId) ||
        // No form designed for this case yet: an automatic one keeps the list usable.
        createDefaultForm(list.title, fields.filter((f) => f.kind !== 'Attachments' || list.enableAttachments));

      const folderPath = props.folderPath && trimSlash(props.folderPath) !== trimSlash(list.rootFolderUrl) ? props.folderPath : undefined;
      if (!cancelled) {
        setLoaded({
          form,
          fields,
          values,
          etag,
          contentTypeId: mode === 'New' ? props.contentTypeId : undefined,
          folderPath,
          listService,
          runtime: {
            mode,
            itemId,
            user: { name: user.name, email: user.email, loginName: user.loginName },
            groups: user.groups
          }
        });
      }
    };
    load().catch((e) => {
      if (!cancelled) {
        setError(errorMessage(e));
      }
    });
    return () => {
      cancelled = true;
    };
  }, [listId, mode, itemId, formId]);

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline>
        The form could not be loaded: {error}
      </MessageBar>
    );
  }
  if (!loaded) {
    return (
      <div className={styles.centered}>
        <Spinner size={SpinnerSize.large} label="Loading form..." />
      </div>
    );
  }
  return (
    <FormView
      form={loaded.form}
      fields={loaded.fields}
      runtime={loaded.runtime}
      listId={listId}
      listService={loaded.listService}
      loadedValues={loaded.values}
      etag={loaded.etag}
      contentTypeId={loaded.contentTypeId}
      folderPath={loaded.folderPath}
      onSaved={props.onSaved}
      onCancel={props.onCancel}
      onEdit={props.onEdit}
    />
  );
};
