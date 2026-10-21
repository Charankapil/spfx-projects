import * as React from 'react';

/**
 * A deliberately small Markdown subset for text controls: **bold**,
 * *italic*, [links](https://...), "- " bullet lines and line breaks.
 * Rendered as React elements, never as HTML, so form text cannot inject
 * markup or script.
 */

function safeHref(url: string): string | undefined {
  const trimmed = url.trim();
  return /^(https?:\/\/|mailto:|tel:|\/)/i.test(trimmed) ? trimmed : undefined;
}

function inline(text: string, keyPrefix: string): React.ReactNode[] {
  const out: React.ReactNode[] = [];
  const pattern = /\*\*([^*]+)\*\*|\*([^*]+)\*|\[([^\]]+)\]\(([^)\s]+)\)/g;
  let last = 0;
  let match: RegExpExecArray | null;
  let i = 0;
  while ((match = pattern.exec(text)) !== null) {
    if (match.index > last) {
      out.push(text.substring(last, match.index));
    }
    const key = `${keyPrefix}-${i++}`;
    if (match[1] !== undefined) {
      out.push(<strong key={key}>{match[1]}</strong>);
    } else if (match[2] !== undefined) {
      out.push(<em key={key}>{match[2]}</em>);
    } else {
      const href = safeHref(match[4]);
      out.push(
        href ? (
          <a key={key} href={href} target="_blank" rel="noopener noreferrer" data-interception="off">
            {match[3]}
          </a>
        ) : (
          match[3]
        )
      );
    }
    last = pattern.lastIndex;
  }
  if (last < text.length) {
    out.push(text.substring(last));
  }
  return out;
}

export function renderRichText(text: string): React.ReactNode {
  const lines = text.split(/\r?\n/);
  const blocks: React.ReactNode[] = [];
  let bullets: string[] = [];
  let paragraph: string[] = [];

  const flushParagraph = (): void => {
    if (paragraph.length) {
      const key = `p${blocks.length}`;
      const children: React.ReactNode[] = [];
      paragraph.forEach((line, i) => {
        if (i > 0) {
          children.push(<br key={`${key}-br${i}`} />);
        }
        children.push(...inline(line, `${key}-${i}`));
      });
      blocks.push(<p key={key}>{children}</p>);
      paragraph = [];
    }
  };
  const flushBullets = (): void => {
    if (bullets.length) {
      const key = `ul${blocks.length}`;
      blocks.push(
        <ul key={key}>
          {bullets.map((b, i) => (
            <li key={i}>{inline(b, `${key}-${i}`)}</li>
          ))}
        </ul>
      );
      bullets = [];
    }
  };

  lines.forEach((line) => {
    const bullet = /^\s*[-*]\s+(.*)$/.exec(line);
    if (bullet) {
      flushParagraph();
      bullets.push(bullet[1]);
    } else if (!line.trim()) {
      flushParagraph();
      flushBullets();
    } else {
      flushBullets();
      paragraph.push(line);
    }
  });
  flushParagraph();
  flushBullets();
  return <>{blocks}</>;
}

const BLOCKED_TAGS = ['script', 'style', 'iframe', 'object', 'embed', 'link', 'meta', 'form', 'input', 'button', 'textarea', 'select'];

/**
 * SharePoint rich text columns hold HTML. SharePoint filters it on save, but
 * values written by other tools may not be, so strip anything active before
 * showing it.
 */
export function sanitizeHtml(html: string): string {
  if (!html) {
    return '';
  }
  const doc = new DOMParser().parseFromString(`<div>${html}</div>`, 'text/html');
  const root = doc.body.firstElementChild as HTMLElement;
  const walk = (el: Element): void => {
    for (let i = el.children.length - 1; i >= 0; i--) {
      const child = el.children[i];
      if (BLOCKED_TAGS.indexOf(child.tagName.toLowerCase()) >= 0) {
        el.removeChild(child);
        continue;
      }
      for (let a = child.attributes.length - 1; a >= 0; a--) {
        const attr = child.attributes[a];
        const name = attr.name.toLowerCase();
        if (name.indexOf('on') === 0 || ((name === 'href' || name === 'src' || name === 'xlink:href') && /^\s*(javascript|data|vbscript):/i.test(attr.value))) {
          child.removeAttribute(attr.name);
        }
      }
      walk(child);
    }
  };
  walk(root);
  return root.innerHTML;
}
