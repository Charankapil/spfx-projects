import * as React from 'react';
import { Icon, SearchBox } from '@fluentui/react';

import { IFieldInfo } from '../../../common/model/IFieldInfo';
import { ControlType } from '../../../common/model/IFormDefinition';
import styles from './Designer.module.scss';

export const DRAG_MIME = 'text/plain';

export type DragPayload =
  | { source: 'field'; field: string }
  | { source: 'control'; type: ControlType }
  | { source: 'move'; controlId: string };

export function setDragData(e: React.DragEvent, payload: DragPayload): void {
  e.dataTransfer.setData(DRAG_MIME, JSON.stringify({ formStudio: payload }));
  e.dataTransfer.effectAllowed = payload.source === 'move' ? 'move' : 'copy';
}

export function getDragData(e: React.DragEvent): DragPayload | undefined {
  try {
    const parsed = JSON.parse(e.dataTransfer.getData(DRAG_MIME));
    return parsed && parsed.formStudio ? (parsed.formStudio as DragPayload) : undefined;
  } catch {
    return undefined;
  }
}

export const FIELD_ICONS: { [kind: string]: string } = {
  Text: 'TextField',
  Note: 'AlignLeft',
  Number: 'NumberField',
  Currency: 'Money',
  Integer: 'NumberField',
  Choice: 'Dropdown',
  MultiChoice: 'CheckList',
  Boolean: 'ToggleRight',
  DateTime: 'Calendar',
  Lookup: 'Link',
  LookupMulti: 'Link',
  User: 'Contact',
  UserMulti: 'People',
  URL: 'Globe',
  TaxonomyFieldType: 'Tag',
  TaxonomyFieldTypeMulti: 'Tag',
  Attachments: 'Attach',
  Calculated: 'Calculator',
  Counter: 'NumberSymbol',
  Unsupported: 'Unknown'
};

export const ELEMENTS: { type: ControlType; label: string; icon: string; help: string }[] = [
  { type: 'heading', label: 'Heading', icon: 'Header1', help: 'A title inside a section' },
  { type: 'text', label: 'Text', icon: 'TextDocument', help: 'Instructions, with **bold**, *italic*, links and bullets' },
  { type: 'image', label: 'Image', icon: 'Photo2', help: 'A logo or illustration' },
  { type: 'divider', label: 'Divider', icon: 'Remove', help: 'A horizontal line' },
  { type: 'spacer', label: 'Spacer', icon: 'CalculatorSubtract', help: 'Empty vertical space' },
  { type: 'button', label: 'Button', icon: 'ButtonControl', help: 'Approve / Reject, set fields and save, open a link, print' },
  { type: 'relatedItems', label: 'Related items', icon: 'Table', help: 'Edit child items of another list inline (order lines, expenses)' }
];

export interface IPaletteProps {
  fields: IFieldInfo[];
  placed: string[];
  onAddField: (field: IFieldInfo) => void;
  onAddElement: (type: ControlType) => void;
}

export const Palette: React.FC<IPaletteProps> = (props) => {
  const [filter, setFilter] = React.useState<string>('');
  const q = filter.toLowerCase();
  const matches = props.fields.filter((f) => !q || f.title.toLowerCase().indexOf(q) >= 0 || f.internalName.toLowerCase().indexOf(q) >= 0);
  const unplaced = matches.filter((f) => props.placed.indexOf(f.internalName) < 0);
  const placed = matches.filter((f) => props.placed.indexOf(f.internalName) >= 0);

  const fieldItem = (f: IFieldInfo, isPlaced: boolean): React.ReactNode => (
    <div
      key={f.internalName}
      className={`${styles.paletteItem} ${isPlaced ? styles.palettePlaced : ''}`}
      draggable
      onDragStart={(e) => setDragData(e, { source: 'field', field: f.internalName })}
      onClick={() => props.onAddField(f)}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          props.onAddField(f);
        }
      }}
      role="button"
      tabIndex={0}
      title={`${f.title} (${f.internalName}, ${f.typeAsString})${f.readOnly ? ' - read-only' : ''}. Drag onto the form, or press Enter to add it to the selected section.`}
    >
      <Icon iconName={FIELD_ICONS[f.kind] || 'TextField'} className={styles.paletteIcon} />
      <span className={styles.paletteLabel}>{f.title}</span>
      {f.required && <span className={styles.requiredDot} title="Required in the list" />}
      {f.readOnly && <Icon iconName="Lock" className={styles.paletteMeta} />}
    </div>
  );

  return (
    <div className={styles.palette}>
      <SearchBox placeholder="Search fields" value={filter} onChange={(_e, v) => setFilter(v || '')} className={styles.paletteSearch} />
      <div className={styles.paletteGroup}>List fields</div>
      {unplaced.map((f) => fieldItem(f, false))}
      {!unplaced.length && <div className={styles.paletteEmpty}>Every field is on the form.</div>}
      {placed.length > 0 && (
        <>
          <div className={styles.paletteGroup}>Already on the form</div>
          {placed.map((f) => fieldItem(f, true))}
        </>
      )}
      <div className={styles.paletteGroup}>Elements</div>
      {ELEMENTS.map((el) => (
        <div
          key={el.type}
          className={styles.paletteItem}
          draggable
          onDragStart={(e) => setDragData(e, { source: 'control', type: el.type })}
          onClick={() => props.onAddElement(el.type)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ' ') {
              e.preventDefault();
              props.onAddElement(el.type);
            }
          }}
          role="button"
          tabIndex={0}
          title={el.help}
        >
          <Icon iconName={el.icon} className={styles.paletteIcon} />
          <span className={styles.paletteLabel}>{el.label}</span>
        </div>
      ))}
    </div>
  );
};
