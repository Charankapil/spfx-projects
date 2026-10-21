import * as React from 'react';
import { Callout, DirectionalHint, IconButton, SearchBox, TextField } from '@fluentui/react';

import { checkExpression, functionCatalog } from '../common/expressions';
import { IFieldInfo } from '../common/model/IFieldInfo';
import styles from './Designer.module.scss';

export interface IExpressionFieldProps {
  label: string;
  value: string | undefined;
  fields: IFieldInfo[];
  placeholder?: string;
  description?: string;
  multiline?: boolean;
  /** Templates: text with {{ expression }} placeholders rather than one expression. */
  template?: boolean;
  onChange: (value: string) => void;
}

const FUNCTIONS = functionCatalog();

function checkTemplate(text: string, known: string[]): { error?: string; unknownFields: string[] } {
  const unknown: string[] = [];
  const pattern = /\{\{([\s\S]+?)\}\}/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    const result = checkExpression(match[1], known);
    if (result.error) {
      return { error: `In {{${match[1].trim()}}}: ${result.error}`, unknownFields: [] };
    }
    result.unknownFields.forEach((f) => unknown.push(f));
  }
  return { unknownFields: unknown };
}

/**
 * Text box for an expression (or a {{ }} template) with live syntax
 * checking and a reference of fields and functions that inserts at the
 * cursor.
 */
export const ExpressionField: React.FC<IExpressionFieldProps> = (props) => {
  const [open, setOpen] = React.useState<boolean>(false);
  const [filter, setFilter] = React.useState<string>('');
  const [draft, setDraft] = React.useState<string>(props.value || '');
  const button = React.useRef<HTMLDivElement>(null);
  // Cursor position, remembered so the reference inserts where the author was typing.
  const caret = React.useRef<{ start: number; end: number } | undefined>();
  const known = React.useMemo(() => props.fields.map((f) => f.internalName), [props.fields]);

  React.useEffect(() => setDraft(props.value || ''), [props.value]);

  const check = props.template ? checkTemplate(draft, known) : checkExpression(draft, known);
  const errorText = check.error || (check.unknownFields.length ? `Unknown field: [${check.unknownFields.join('], [')}]. Use internal names.` : undefined);

  const commit = (value: string): void => {
    setDraft(value);
    props.onChange(value);
  };

  const insert = (text: string): void => {
    const snippet = props.template && !/\{\{/.test(text) ? `{{${text}}}` : text;
    const pos = caret.current;
    if (pos && pos.end <= draft.length) {
      commit(draft.substring(0, pos.start) + snippet + draft.substring(pos.end));
    } else {
      commit(draft + snippet);
    }
    setOpen(false);
  };

  const q = filter.toLowerCase();
  const fieldMatches = props.fields.filter((f) => !q || f.title.toLowerCase().indexOf(q) >= 0 || f.internalName.toLowerCase().indexOf(q) >= 0);
  const fnMatches = FUNCTIONS.filter((f) => !q || f.name.toLowerCase().indexOf(q) >= 0 || f.help.toLowerCase().indexOf(q) >= 0);

  return (
    <div className={styles.expression}>
      <div className={styles.expressionRow}>
        <TextField
          label={props.label}
          value={draft}
          placeholder={props.placeholder}
          multiline={props.multiline}
          autoAdjustHeight={props.multiline}
          className={styles.code}
          onSelect={(e) => {
            const el = e.target as HTMLInputElement;
            caret.current = { start: el.selectionStart || 0, end: el.selectionEnd || el.selectionStart || 0 };
          }}
          onChange={(_e, v) => setDraft(v || '')}
          onBlur={() => {
            if (draft !== (props.value || '')) {
              props.onChange(draft);
            }
          }}
          errorMessage={draft.trim() ? errorText : undefined}
          description={props.description}
        />
        <div ref={button} className={styles.expressionHelp}>
          <IconButton iconProps={{ iconName: 'Variable2' }} title="Insert a field or function" ariaLabel="Insert a field or function" onClick={() => setOpen(!open)} />
        </div>
      </div>
      {open && (
        <Callout target={button.current} onDismiss={() => setOpen(false)} directionalHint={DirectionalHint.leftTopEdge} setInitialFocus className={styles.referenceCallout}>
          <div className={styles.reference}>
            <SearchBox placeholder="Search fields and functions" value={filter} onChange={(_e, v) => setFilter(v || '')} />
            <div className={styles.referenceGroup}>Fields</div>
            {fieldMatches.map((f) => (
              <button key={f.internalName} type="button" className={styles.referenceItem} onClick={() => insert(`[${f.internalName}]`)}>
                <span className={styles.referenceName}>{f.title}</span>
                <span className={styles.referenceMeta}>[{f.internalName}] · {f.kind}</span>
              </button>
            ))}
            <div className={styles.referenceGroup}>Functions</div>
            {fnMatches.map((f) => (
              <button key={f.name} type="button" className={styles.referenceItem} onClick={() => insert(`${f.name}(`)}>
                <span className={styles.referenceName}>{f.signature}</span>
                <span className={styles.referenceMeta}>{f.help}</span>
              </button>
            ))}
            <div className={styles.referenceGroup}>Operators</div>
            <div className={styles.referenceMeta} style={{ padding: '0 8px 8px' }}>
              = &lt;&gt; &lt; &gt; &lt;= &gt;= · And Or Not (&amp;&amp; || !) · + - * / · &amp; joins text · &quot;x&quot; in [Choices] · [Person].Email · [Lookup].Id
            </div>
          </div>
        </Callout>
      )}
    </div>
  );
};
