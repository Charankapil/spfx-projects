import * as React from 'react';
import { IconButton } from '@fluentui/react';

import { sanitizeHtml } from '../../richContent';
import styles from '../FormRenderer.module.scss';

export interface IRichTextEditorProps {
  id: string;
  label: string;
  value: string;
  rows: number;
  onChange: (html: string) => void;
}

const COMMANDS: { command: string; icon: string; title: string }[] = [
  { command: 'bold', icon: 'Bold', title: 'Bold' },
  { command: 'italic', icon: 'Italic', title: 'Italic' },
  { command: 'underline', icon: 'Underline', title: 'Underline' },
  { command: 'insertUnorderedList', icon: 'BulletedList', title: 'Bulleted list' },
  { command: 'insertOrderedList', icon: 'NumberedList', title: 'Numbered list' }
];

/** Minimal rich text editor for enhanced rich text columns. */
export const RichTextEditor: React.FC<IRichTextEditorProps> = (props) => {
  const ref = React.useRef<HTMLDivElement>(null);
  const lastValue = React.useRef<string>('');

  // Only push value changes from outside (defaults, calculations); typing is not re-rendered.
  React.useEffect(() => {
    if (ref.current && props.value !== lastValue.current) {
      ref.current.innerHTML = sanitizeHtml(props.value || '');
      lastValue.current = props.value || '';
    }
  }, [props.value]);

  const emit = (): void => {
    if (ref.current) {
      const html = ref.current.innerHTML;
      lastValue.current = html;
      props.onChange(html === '<br>' ? '' : html);
    }
  };

  return (
    <div className={styles.richText}>
      <div className={styles.richToolbar} role="toolbar" aria-label={`${props.label} formatting`}>
        {COMMANDS.map((c) => (
          <IconButton
            key={c.command}
            iconProps={{ iconName: c.icon }}
            title={c.title}
            ariaLabel={c.title}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              document.execCommand(c.command);
              emit();
            }}
          />
        ))}
      </div>
      <div
        id={props.id}
        ref={ref}
        className={styles.richBody}
        style={{ minHeight: props.rows * 20 }}
        contentEditable
        role="textbox"
        aria-multiline="true"
        aria-label={props.label}
        onInput={emit}
        onBlur={emit}
      />
    </div>
  );
};
