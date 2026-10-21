import * as React from 'react';
import { Icon, MessageBar, MessageBarType, SearchBox, Spinner, SpinnerSize } from '@fluentui/react';

import { errorMessage, ISpContext } from '../../../common/services/http';
import { IListSummary, ListService } from '../../../common/services/ListService';
import styles from './Designer.module.scss';
import { ListWorkspace } from './ListWorkspace';

export interface IFormStudioDesignerProps {
  ctx: ISpContext;
  listId: string;
  onListChange: (listId: string) => void;
}

/** Entry point: pick a list, then design its forms. */
export const FormStudioDesigner: React.FC<IFormStudioDesignerProps> = (props) => {
  const [listId, setListId] = React.useState<string>(props.listId);
  const [lists, setLists] = React.useState<IListSummary[] | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [filter, setFilter] = React.useState<string>('');

  React.useEffect(() => setListId(props.listId), [props.listId]);

  React.useEffect(() => {
    if (listId || lists) {
      return;
    }
    new ListService(props.ctx)
      .getLists()
      .then(setLists)
      .catch((e) => setError(errorMessage(e)));
  }, [listId]);

  const choose = (id: string): void => {
    setListId(id);
    props.onListChange(id);
  };

  if (listId) {
    return <ListWorkspace key={listId} ctx={props.ctx} listId={listId} onChangeList={() => choose('')} />;
  }

  const q = filter.toLowerCase();
  const shown = (lists || []).filter((l) => !q || l.title.toLowerCase().indexOf(q) >= 0);

  return (
    <div className={styles.picker}>
      <div className={styles.pickerHeader}>
        <Icon iconName="FormLibrary" className={styles.pickerIcon} />
        <div>
          <h2 className={styles.pickerTitle}>Form Studio</h2>
          <div className={styles.help}>Design custom forms for a SharePoint list - no Power Apps needed. Choose a list on this site to start.</div>
        </div>
      </div>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {!lists && !error && <Spinner size={SpinnerSize.medium} label="Loading lists..." />}
      {lists && (
        <>
          <SearchBox placeholder="Find a list" value={filter} onChange={(_e, v) => setFilter(v || '')} className={styles.pickerSearch} />
          <div className={styles.listGrid}>
            {shown.map((l) => (
              <button key={l.id} type="button" className={styles.listCard} onClick={() => choose(l.id)}>
                <Icon iconName={l.isLibrary ? 'DocLibrary' : 'BulletedList2'} className={styles.listCardIcon} />
                <span className={styles.listCardTitle}>{l.title}</span>
                <span className={styles.listCardMeta}>
                  {l.isLibrary ? 'Library' : 'List'} · {l.itemCount} item{l.itemCount === 1 ? '' : 's'}
                </span>
              </button>
            ))}
            {!shown.length && <div className={styles.help}>No lists match.</div>}
          </div>
        </>
      )}
    </div>
  );
};
