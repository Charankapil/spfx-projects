import * as React from 'react';
import { DefaultButton, Icon, IconButton } from '@fluentui/react';

import { createControl, newId } from '../common/model/formFactory';
import { IFieldInfo } from '../common/model/IFieldInfo';
import { IControl, IFormDefinition, ISection } from '../common/model/IFormDefinition';
import styles from './Designer.module.scss';
import {
  addControl,
  addPage,
  addSection,
  findControl,
  moveControl,
  nudgeControl,
  nudgeSection,
  removeControl,
  removeSection,
  Selection
} from './designerOps';
import { ELEMENTS, FIELD_ICONS, getDragData, setDragData } from './Palette';

export interface ICanvasProps {
  form: IFormDefinition;
  fields: IFieldInfo[];
  selection: Selection;
  activePageId: string;
  onActivePage: (pageId: string) => void;
  onSelect: (selection: Selection) => void;
  onChange: (form: IFormDefinition, selection?: Selection) => void;
}

interface IDropTarget {
  sectionId: string;
  index: number;
}

function controlTitle(control: IControl, fields: IFieldInfo[]): { title: string; meta: string; icon: string; missing?: boolean } {
  if (control.type === 'field') {
    const field = fields.filter((f) => f.internalName === control.field)[0];
    if (!field) {
      return { title: control.label || control.field || '?', meta: 'Field no longer exists in the list', icon: 'Warning', missing: true };
    }
    return { title: control.label || field.title, meta: field.typeAsString, icon: FIELD_ICONS[field.kind] || 'TextField' };
  }
  const element = ELEMENTS.filter((e) => e.type === control.type)[0];
  const title =
    control.type === 'heading' || control.type === 'text'
      ? (control.content || '').split('\n')[0].substring(0, 80) || element.label
      : control.type === 'button' || control.type === 'relatedItems'
      ? control.label || element.label
      : element.label;
  return { title, meta: element.label, icon: element.icon };
}

function badges(control: IControl, fields: IFieldInfo[]): string[] {
  const out: string[] = [];
  const field = fields.filter((f) => f.internalName === control.field)[0];
  if (field && field.required) {
    out.push('Required');
  } else if (control.requiredWhen) {
    out.push('Required if');
  }
  if (control.visibleWhen) {
    out.push('Visible if');
  }
  if (control.readOnlyWhen) {
    out.push('Read-only if');
  }
  if (control.calculatedValue) {
    out.push('Calculated');
  }
  if (control.defaultValue) {
    out.push('Default');
  }
  if (control.validations && control.validations.length) {
    out.push(`${control.validations.length} rule${control.validations.length > 1 ? 's' : ''}`);
  }
  if (control.lookupCascade || control.lookupFilter) {
    out.push('Filtered');
  }
  if (control.assignments && control.assignments.length) {
    out.push(`Sets ${control.assignments.length} field${control.assignments.length > 1 ? 's' : ''}`);
  }
  return out;
}

/** The design surface: pages, sections and controls with drag-and-drop. */
export const Canvas: React.FC<ICanvasProps> = (props) => {
  const { form, fields, selection, onChange, onSelect } = props;
  const [drop, setDrop] = React.useState<IDropTarget | undefined>();
  const page = form.pages.filter((p) => p.id === props.activePageId)[0] || form.pages[0];

  const handleDrop = (e: React.DragEvent, sectionId: string, index: number): void => {
    e.preventDefault();
    e.stopPropagation();
    setDrop(undefined);
    const data = getDragData(e);
    if (!data) {
      return;
    }
    if (data.source === 'move') {
      onChange(moveControl(form, data.controlId, sectionId, index), { kind: 'control', id: data.controlId });
      return;
    }
    if (data.source === 'field') {
      // A field appears once per form: dropping it again moves the existing control.
      const existing = form.pages
        .map((p) => p.sections.map((s) => s.controls.filter((c) => c.type === 'field' && c.field === data.field)[0]).filter((c) => !!c)[0])
        .filter((c) => !!c)[0];
      if (existing) {
        onChange(moveControl(form, existing.id, sectionId, index), { kind: 'control', id: existing.id });
        return;
      }
    }
    const control =
      data.source === 'field'
        ? createControl('field', fields.filter((f) => f.internalName === data.field)[0])
        : createControl(data.type);
    onChange(addControl(form, sectionId, control, index), { kind: 'control', id: control.id });
  };

  const dragOver = (e: React.DragEvent, sectionId: string, index: number): void => {
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = e.dataTransfer.effectAllowed === 'move' ? 'move' : 'copy';
    if (!drop || drop.sectionId !== sectionId || drop.index !== index) {
      setDrop({ sectionId, index });
    }
  };

  const renderControl = (control: IControl, section: ISection, index: number): React.ReactNode => {
    const info = controlTitle(control, fields);
    const selected = selection.kind === 'control' && selection.id === control.id;
    const span = Math.max(1, Math.min(control.span || 1, section.columns));
    const showIndicator = drop && drop.sectionId === section.id && drop.index === index;
    return (
      <div
        key={control.id}
        className={`${styles.canvasControl} ${selected ? styles.selected : ''} ${info.missing ? styles.missing : ''} ${showIndicator ? styles.dropBefore : ''}`}
        style={{ gridColumn: `span ${span}` }}
        draggable
        onDragStart={(e) => {
          e.stopPropagation();
          setDragData(e, { source: 'move', controlId: control.id });
        }}
        onDragEnd={() => setDrop(undefined)}
        onDragOver={(e) => dragOver(e, section.id, index)}
        onDrop={(e) => handleDrop(e, section.id, index)}
        onClick={(e) => {
          e.stopPropagation();
          onSelect({ kind: 'control', id: control.id });
        }}
        onKeyDown={(e) => {
          if (e.target !== e.currentTarget) {
            return;
          }
          if (e.key === 'Delete') {
            onChange(removeControl(form, control.id), { kind: 'section', id: section.id });
          } else if (e.key === 'Enter') {
            onSelect({ kind: 'control', id: control.id });
          }
        }}
        role="button"
        tabIndex={0}
        aria-label={`${info.title}, ${info.meta}`}
        aria-pressed={selected}
      >
        <div className={styles.controlHeader}>
          <Icon iconName={info.icon} className={styles.controlIcon} />
          <div className={styles.controlText}>
            <div className={styles.controlTitle}>{info.title}</div>
            <div className={styles.controlMeta}>{info.meta}</div>
          </div>
          <div className={styles.controlTools}>
            <IconButton iconProps={{ iconName: 'Up' }} title="Move earlier" ariaLabel="Move earlier" disabled={index === 0} onClick={(e) => { e.stopPropagation(); onChange(nudgeControl(form, control.id, -1)); }} />
            <IconButton iconProps={{ iconName: 'Down' }} title="Move later" ariaLabel="Move later" disabled={index === section.controls.length - 1} onClick={(e) => { e.stopPropagation(); onChange(nudgeControl(form, control.id, 1)); }} />
            {control.type !== 'field' && (
              <IconButton
                iconProps={{ iconName: 'Copy' }}
                title="Duplicate"
                ariaLabel="Duplicate"
                onClick={(e) => {
                  e.stopPropagation();
                  const copy: IControl = { ...JSON.parse(JSON.stringify(control)), id: newId('c') };
                  onChange(addControl(form, section.id, copy, index + 1), { kind: 'control', id: copy.id });
                }}
              />
            )}
            <IconButton iconProps={{ iconName: 'Delete' }} title="Remove from form" ariaLabel="Remove from form" onClick={(e) => { e.stopPropagation(); onChange(removeControl(form, control.id), { kind: 'section', id: section.id }); }} />
          </div>
        </div>
        {control.type === 'field' && <div className={styles.mockInput} />}
        {badges(control, fields).length > 0 && (
          <div className={styles.badges}>
            {badges(control, fields).map((b) => (
              <span key={b} className={styles.badge}>
                {b}
              </span>
            ))}
          </div>
        )}
      </div>
    );
  };

  const renderSection = (section: ISection, sectionIndex: number): React.ReactNode => {
    const selected = selection.kind === 'section' && selection.id === section.id;
    const containsSelection = selection.kind === 'control' && section.controls.some((c) => c.id === selection.id);
    const endIndicator = drop && drop.sectionId === section.id && drop.index === section.controls.length;
    return (
      <div
        key={section.id}
        className={`${styles.canvasSection} ${selected ? styles.selected : ''} ${containsSelection ? styles.sectionActive : ''}`}
        onClick={() => onSelect({ kind: 'section', id: section.id })}
      >
        <div className={styles.sectionHeader}>
          <button type="button" className={styles.sectionName} onClick={(e) => { e.stopPropagation(); onSelect({ kind: 'section', id: section.id }); }}>
            {section.title || <span className={styles.untitled}>Section {sectionIndex + 1}</span>}
            <span className={styles.sectionMeta}>
              {section.columns} column{section.columns > 1 ? 's' : ''}
              {section.collapsible ? ' · collapsible' : ''}
              {section.visibleWhen ? ' · conditional' : ''}
            </span>
          </button>
          <div className={styles.controlTools}>
            <IconButton iconProps={{ iconName: 'Up' }} title="Move section up" ariaLabel="Move section up" onClick={(e) => { e.stopPropagation(); onChange(nudgeSection(form, section.id, -1)); }} />
            <IconButton iconProps={{ iconName: 'Down' }} title="Move section down" ariaLabel="Move section down" onClick={(e) => { e.stopPropagation(); onChange(nudgeSection(form, section.id, 1)); }} />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete section"
              ariaLabel="Delete section"
              onClick={(e) => {
                e.stopPropagation();
                // eslint-disable-next-line no-alert
                if (!section.controls.length || window.confirm('Delete this section and everything in it?')) {
                  onChange(removeSection(form, section.id), { kind: 'page', id: page.id });
                }
              }}
            />
          </div>
        </div>
        <div
          className={`${styles.canvasGrid} ${endIndicator ? styles.dropEnd : ''}`}
          style={{ gridTemplateColumns: `repeat(${section.columns}, minmax(0, 1fr))` }}
          onDragOver={(e) => dragOver(e, section.id, section.controls.length)}
          onDragLeave={(e) => {
            if (!e.currentTarget.contains(e.relatedTarget as Node)) {
              setDrop(undefined);
            }
          }}
          onDrop={(e) => handleDrop(e, section.id, section.controls.length)}
        >
          {section.controls.map((c, i) => renderControl(c, section, i))}
          {!section.controls.length && <div className={styles.emptySection}>Drag fields or elements here</div>}
        </div>
        <div className={styles.addSectionRow}>
          <DefaultButton
            iconProps={{ iconName: 'Add' }}
            text="Add section"
            className={styles.addSectionButton}
            onClick={(e) => {
              e.stopPropagation();
              const result = addSection(form, page.id, section.id);
              onChange(result.form, { kind: 'section', id: result.sectionId });
            }}
          />
        </div>
      </div>
    );
  };

  // Keep the selected control's page visible.
  React.useEffect(() => {
    if (selection.kind === 'control') {
      const loc = findControl(form, selection.id);
      if (loc && loc.page.id !== page.id) {
        props.onActivePage(loc.page.id);
      }
    }
  }, [selection]);

  return (
    <div className={styles.canvas} onClick={() => onSelect({ kind: 'form' })}>
      <div className={styles.pageTabs} role="tablist">
        {form.pages.map((p) => (
          <button
            key={p.id}
            type="button"
            role="tab"
            aria-selected={p.id === page.id}
            className={`${styles.pageTab} ${p.id === page.id ? styles.pageTabActive : ''}`}
            onClick={(e) => {
              e.stopPropagation();
              props.onActivePage(p.id);
              onSelect({ kind: 'page', id: p.id });
            }}
            // Dropping on a tab puts the control at the end of that page.
            onDragOver={(e) => {
              e.preventDefault();
              if (p.id !== page.id) {
                props.onActivePage(p.id);
              }
            }}
            onDrop={(e) => {
              const last = p.sections[p.sections.length - 1];
              handleDrop(e, last.id, last.controls.length);
            }}
          >
            {p.title || 'Untitled page'}
            {p.visibleWhen && <Icon iconName="Filter" className={styles.pageTabIcon} />}
          </button>
        ))}
        <button
          type="button"
          className={styles.pageTab}
          title={form.layout === 'sections' ? 'Adding a page switches the form to tabs' : 'Add page'}
          onClick={(e) => {
            e.stopPropagation();
            const result = addPage(form);
            onChange(result.form, { kind: 'page', id: result.pageId });
            props.onActivePage(result.pageId);
          }}
        >
          <Icon iconName="Add" /> {form.layout === 'wizard' ? 'Step' : 'Page'}
        </button>
        <span className={styles.layoutHint}>
          {form.layout === 'wizard' ? 'Wizard: one page per step' : form.layout === 'tabs' ? 'Tabs' : 'Single page'}
        </span>
      </div>
      {page.sections.map(renderSection)}
    </div>
  );
};
