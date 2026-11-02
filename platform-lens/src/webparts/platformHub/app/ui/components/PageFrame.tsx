import * as React from 'react';
import { Link, Panel, PanelType, Text } from '@fluentui/react';
import { IDetail, PageId, useApp } from '../AppContext';
import { pageDef } from '../exportHelpers';
import { ExportMenu } from './ExportMenu';
import { PageHeader } from './Layout';
import { safeUrl } from '../../logic/format';

/** Standard page chrome: title, description, page actions and the Export menu. */
export const PageFrame: React.FC<{ id: PageId; actions?: React.ReactNode; children?: React.ReactNode }> = (props) => {
  const ref = React.useRef<HTMLDivElement>(null);
  const def = pageDef(props.id);
  return (
    <div ref={ref}>
      <PageHeader
        title={def ? def.title : props.id}
        description={def ? def.description : undefined}
        actions={
          <>
            {props.actions}
            {def && def.build && <ExportMenu page={props.id} rootRef={ref} />}
          </>
        }
      />
      {props.children}
    </div>
  );
};

export const DetailPanel: React.FC<{ detail: IDetail | undefined; onDismiss: () => void }> = ({ detail, onDismiss }) => {
  const { theme } = useApp();
  return (
    <Panel
      isOpen={!!detail}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={detail ? detail.title : ''}
      closeButtonAriaLabel="Close"
      isLightDismiss
    >
      {detail && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14, paddingBottom: 24 }}>
          {detail.subtitle && <Text style={{ color: theme.semanticColors.bodySubtext }}>{detail.subtitle}</Text>}
          <dl style={{ display: 'grid', gridTemplateColumns: 'minmax(110px, 34%) 1fr', gap: '6px 12px', margin: 0 }}>
            {detail.fields.map((f) => (
              <React.Fragment key={f.label}>
                <dt style={{ color: theme.semanticColors.bodySubtext, fontSize: 12 }}>{f.label}</dt>
                <dd style={{ margin: 0, fontSize: 13, wordBreak: 'break-word' }}>{f.value === undefined || f.value === '' ? '—' : f.value}</dd>
              </React.Fragment>
            ))}
          </dl>
          {detail.links && detail.links.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              {detail.links.filter((l) => !!safeUrl(l.href)).map((l) => (
                <Link key={l.href} href={safeUrl(l.href)} target="_blank" rel="noopener noreferrer">{l.label} ↗</Link>
              ))}
            </div>
          )}
          {detail.extra}
        </div>
      )}
    </Panel>
  );
};
