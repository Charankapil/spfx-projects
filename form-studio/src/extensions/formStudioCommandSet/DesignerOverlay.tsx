import * as React from 'react';
import { FocusTrapZone, Icon, IconButton } from '@fluentui/react';

import { ISpContext } from '../../common/services/http';
import { ListWorkspace } from '../../designer/ListWorkspace';
import styles from './DesignerOverlay.module.scss';

export interface IDesignerOverlayProps {
  ctx: ISpContext;
  listId: string;
  listTitle: string;
  onClose: () => void;
}

/** The designer as a full-screen window over the list. */
export const DesignerOverlay: React.FC<IDesignerOverlayProps> = (props) => {
  const dirty = React.useRef<boolean>(false);
  const root = React.useRef<HTMLDivElement>(null);

  const close = (): void => {
    // eslint-disable-next-line no-alert
    if (dirty.current && !window.confirm('You have unsaved form changes. Close the designer and lose them?')) {
      return;
    }
    props.onClose();
  };

  // The list page behind keeps scrolling otherwise.
  React.useEffect(() => {
    const previous = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.body.style.overflow = previous;
    };
  }, []);

  return (
    <FocusTrapZone isClickableOutsideFocusTrap={false} forceFocusInsideTrap={false}>
      <div
        ref={root}
        className={styles.overlay}
        role="dialog"
        aria-modal="true"
        aria-label={`Form Studio - ${props.listTitle}`}
        onKeyDown={(e) => {
          // Drop-downs, callouts and dialogs render in portals outside this element;
          // their Escape still bubbles here through React, so ignore it - it closes them.
          if (e.key === 'Escape' && !e.defaultPrevented && root.current && root.current.contains(e.target as Node)) {
            close();
          }
        }}
        tabIndex={-1}
      >
        <div className={styles.titleBar}>
          <Icon iconName="FormLibrary" className={styles.icon} />
          <span className={styles.title}>Form Studio</span>
          <span className={styles.hint}>Design the forms people use to add, edit and view items in this list.</span>
          <IconButton className={styles.close} iconProps={{ iconName: 'Cancel' }} title="Close" ariaLabel="Close Form Studio" onClick={close} />
        </div>
        <div className={styles.body}>
          <ListWorkspace ctx={props.ctx} listId={props.listId} fill onDirtyChange={(d) => (dirty.current = d)} />
        </div>
      </div>
    </FocusTrapZone>
  );
};
