import * as React from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Pivot,
  PivotItem,
  PrimaryButton,
  TextField
} from '@fluentui/react';
import styles from './ESign.module.scss';

/**
 * Adopt a signature or initials: draw it, type it in a handwriting font, or
 * upload an image. Whatever the method, the result is a trimmed PNG data
 * URL, which is what gets recorded and stamped into the PDF.
 */

export interface ISignatureDialogProps {
  kind: 'signature' | 'initials';
  defaultText: string;
  onAdopt: (dataUrl: string, remember: boolean) => void;
  onDismiss: () => void;
}

const FONTS = ["'Segoe Script', cursive", "'Brush Script MT', cursive", "'Lucida Handwriting', cursive", "'Ink Free', 'Bradley Hand', cursive"];
const INK = '#0b1f5c';
const MAX_W = 480;
const MAX_H = 160;
/** Keep each image well inside what a SharePoint multi-line text column holds (~64K characters). */
const MAX_DATA_URL = 24000;

/** Crop transparent margins and scale down so the stored image stays small. */
function trimmed(source: HTMLCanvasElement): string | undefined {
  const context = source.getContext('2d') as CanvasRenderingContext2D;
  const { width, height } = source;
  const data = context.getImageData(0, 0, width, height).data;
  let top = height;
  let left = width;
  let right = -1;
  let bottom = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 8) {
        top = Math.min(top, y);
        bottom = Math.max(bottom, y);
        left = Math.min(left, x);
        right = Math.max(right, x);
      }
    }
  }
  if (right < 0) {
    return undefined;
  }
  const pad = 6;
  const w = right - left + 1 + pad * 2;
  const h = bottom - top + 1 + pad * 2;
  let scale = Math.min(1, MAX_W / w, MAX_H / h);
  for (;;) {
    const out = document.createElement('canvas');
    out.width = Math.max(1, Math.round(w * scale));
    out.height = Math.max(1, Math.round(h * scale));
    (out.getContext('2d') as CanvasRenderingContext2D).drawImage(source, left - pad, top - pad, w, h, 0, 0, out.width, out.height);
    const url = out.toDataURL('image/png');
    // Noisy scans compress badly; shrink until the image fits.
    if (url.length <= MAX_DATA_URL || out.width < 120) {
      return url;
    }
    scale *= 0.8;
  }
}

function renderTyped(text: string, font: string): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = 900;
  canvas.height = 220;
  const context = canvas.getContext('2d') as CanvasRenderingContext2D;
  context.fillStyle = INK;
  context.textBaseline = 'middle';
  let size = 110;
  context.font = `${size}px ${font}`;
  while (size > 20 && context.measureText(text).width > canvas.width - 40) {
    size -= 4;
    context.font = `${size}px ${font}`;
  }
  context.fillText(text, 20, canvas.height / 2);
  return canvas;
}

export function SignatureDialog(props: ISignatureDialogProps): JSX.Element {
  const [tab, setTab] = React.useState('draw');
  const [text, setText] = React.useState(props.defaultText);
  const [font, setFont] = React.useState(FONTS[0]);
  const [upload, setUpload] = React.useState<string>();
  const [hasInk, setHasInk] = React.useState(false);
  const [remember, setRemember] = React.useState(true);
  const canvas = React.useRef<HTMLCanvasElement | null>(null);
  const drawing = React.useRef<{ x: number; y: number } | undefined>();

  const setupCanvas = React.useCallback((element: HTMLCanvasElement | null) => {
    canvas.current = element;
    if (element) {
      const ratio = window.devicePixelRatio || 1;
      element.width = element.clientWidth * ratio;
      element.height = element.clientHeight * ratio;
      const context = element.getContext('2d') as CanvasRenderingContext2D;
      context.scale(ratio, ratio);
      context.lineCap = 'round';
      context.lineJoin = 'round';
      context.strokeStyle = INK;
      context.lineWidth = 2.6;
      // A fresh canvas (e.g. after switching tabs) holds no ink.
      setHasInk(false);
    }
  }, []);

  const point = (e: React.PointerEvent<HTMLCanvasElement>): { x: number; y: number } => {
    const rect = e.currentTarget.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const onDown = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    e.currentTarget.setPointerCapture(e.pointerId);
    drawing.current = point(e);
    const context = e.currentTarget.getContext('2d') as CanvasRenderingContext2D;
    context.beginPath();
    context.arc(drawing.current.x, drawing.current.y, 1.2, 0, Math.PI * 2);
    context.fillStyle = INK;
    context.fill();
    setHasInk(true);
  };

  const onMove = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    if (!drawing.current) {
      return;
    }
    const next = point(e);
    const context = e.currentTarget.getContext('2d') as CanvasRenderingContext2D;
    // Thinner strokes when moving fast, like a real pen.
    const speed = Math.hypot(next.x - drawing.current.x, next.y - drawing.current.y);
    context.lineWidth = Math.max(1.4, 3.2 - speed / 12);
    context.beginPath();
    context.moveTo(drawing.current.x, drawing.current.y);
    context.lineTo(next.x, next.y);
    context.stroke();
    drawing.current = next;
  };

  const clear = (): void => {
    const element = canvas.current;
    if (element) {
      (element.getContext('2d') as CanvasRenderingContext2D).clearRect(0, 0, element.width, element.height);
    }
    setHasInk(false);
  };

  const onFile = (file: File | undefined): void => {
    if (!file) {
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const image = new Image();
      image.onload = () => {
        const c = document.createElement('canvas');
        const scale = Math.min(1, 900 / image.width, 300 / image.height);
        c.width = image.width * scale;
        c.height = image.height * scale;
        const context = c.getContext('2d') as CanvasRenderingContext2D;
        context.drawImage(image, 0, 0, c.width, c.height);
        // Make a white background transparent so scans sit cleanly on the page.
        const pixels = context.getImageData(0, 0, c.width, c.height);
        for (let i = 0; i < pixels.data.length; i += 4) {
          if (pixels.data[i] > 235 && pixels.data[i + 1] > 235 && pixels.data[i + 2] > 235) {
            pixels.data[i + 3] = 0;
          }
        }
        context.putImageData(pixels, 0, 0);
        setUpload(trimmed(c));
      };
      image.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  };

  const result = (): string | undefined => {
    if (tab === 'draw') {
      return hasInk && canvas.current ? trimmed(canvas.current) : undefined;
    }
    if (tab === 'type') {
      return text.trim() ? trimmed(renderTyped(text.trim(), font)) : undefined;
    }
    return upload;
  };

  const ready = tab === 'draw' ? hasInk : tab === 'type' ? !!text.trim() : !!upload;
  const noun = props.kind === 'signature' ? 'signature' : 'initials';

  return (
    <Dialog
      hidden={false}
      onDismiss={props.onDismiss}
      minWidth={Math.min(560, window.innerWidth - 32)}
      dialogContentProps={{ type: DialogType.normal, title: `Adopt your ${noun}` }}
      modalProps={{ isBlocking: true }}
    >
      <Pivot selectedKey={tab} onLinkClick={item => setTab(item?.props.itemKey || 'draw')}>
        <PivotItem headerText="Draw" itemKey="draw" itemIcon="EditCreate">
          <p className={styles.muted}>Draw your {noun} with a mouse, pen or finger.</p>
          <canvas
            ref={setupCanvas}
            className={styles.padCanvas}
            onPointerDown={onDown}
            onPointerMove={onMove}
            onPointerUp={() => (drawing.current = undefined)}
            onPointerLeave={() => (drawing.current = undefined)}
            aria-label={`Draw your ${noun}`}
          />
          <DefaultButton text="Clear" onClick={clear} styles={{ root: { marginTop: 8 } }} />
        </PivotItem>
        <PivotItem headerText="Type" itemKey="type" itemIcon="Font">
          <TextField label={props.kind === 'signature' ? 'Full name' : 'Initials'} value={text} onChange={(_, v) => setText(v || '')} />
          <div className={styles.fonts}>
            {FONTS.map(f => (
              <div
                key={f}
                role="button"
                tabIndex={0}
                className={`${styles.typed} ${f === font ? styles.typedSelected : ''}`}
                style={{ fontFamily: f, color: INK, fontSize: props.kind === 'signature' ? 28 : 36 }}
                onClick={() => setFont(f)}
                onKeyDown={e => (e.key === 'Enter' || e.key === ' ') && setFont(f)}
              >
                {text || 'Your name'}
              </div>
            ))}
          </div>
        </PivotItem>
        <PivotItem headerText="Upload" itemKey="upload" itemIcon="Upload">
          <p className={styles.muted}>Upload a photo or scan of your {noun} on white paper (PNG or JPG).</p>
          <input type="file" accept="image/png,image/jpeg" onChange={e => onFile(e.target.files?.[0])} />
          {upload && (
            <div className={styles.typed} style={{ marginTop: 8 }}>
              <img src={upload} alt="Uploaded signature" className={styles.fieldImage} />
            </div>
          )}
        </PivotItem>
      </Pivot>
      <Checkbox
        styles={{ root: { marginTop: 16 } }}
        label={`Remember my ${noun} on this device`}
        checked={remember}
        onChange={(_, v) => setRemember(!!v)}
      />
      <p className={styles.muted}>
        By selecting Adopt and sign, I agree that this {noun} will be the electronic representation of my {noun} for all purposes when I use
        it on documents.
      </p>
      <DialogFooter>
        <PrimaryButton
          text="Adopt and sign"
          disabled={!ready}
          onClick={() => {
            const value = result();
            if (value) {
              props.onAdopt(value, remember);
            }
          }}
        />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}

const STORAGE_KEY = 'esign.adopted.';

/** Saved signature for this user on this device (localStorage can be unavailable - never required). */
export function savedSignature(kind: 'signature' | 'initials', email: string): string | undefined {
  try {
    return window.localStorage.getItem(`${STORAGE_KEY}${kind}.${email.toLowerCase()}`) || undefined;
  } catch {
    return undefined;
  }
}

export function rememberSignature(kind: 'signature' | 'initials', email: string, value: string | undefined): void {
  try {
    const key = `${STORAGE_KEY}${kind}.${email.toLowerCase()}`;
    if (value) {
      window.localStorage.setItem(key, value);
    } else {
      window.localStorage.removeItem(key);
    }
  } catch {
    // Private mode or blocked storage: the signature just is not remembered.
  }
}
