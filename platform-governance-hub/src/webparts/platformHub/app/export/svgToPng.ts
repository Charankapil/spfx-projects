import { downloadBlob, fileStamp, safeFileName } from './tabular';

/**
 * Rasterises an inline SVG chart to a PNG data URL. Charts set all colours and
 * fonts as attributes, so the serialised SVG renders the same outside the page.
 */
export function svgToPngDataUrl(svg: SVGSVGElement, background: string, scale: number = 2): Promise<{ dataUrl: string; width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const box = svg.getBoundingClientRect();
    const width = Math.max(1, Math.round(box.width || Number(svg.getAttribute('width')) || 600));
    const height = Math.max(1, Math.round(box.height || Number(svg.getAttribute('height')) || 300));
    const clone = svg.cloneNode(true) as SVGSVGElement;
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    clone.setAttribute('width', String(width));
    clone.setAttribute('height', String(height));
    const xml = new XMLSerializer().serializeToString(clone);
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = width * scale;
        canvas.height = height * scale;
        const ctx = canvas.getContext('2d');
        if (!ctx) throw new Error('Canvas is not available');
        ctx.fillStyle = background;
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, width, height);
        resolve({ dataUrl: canvas.toDataURL('image/png'), width, height });
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = () => reject(new Error('Could not render the chart image'));
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(xml);
  });
}

export async function downloadSvgAsPng(svg: SVGSVGElement, title: string, background: string): Promise<void> {
  const { dataUrl } = await svgToPngDataUrl(svg, background);
  const bin = atob(dataUrl.split(',')[1]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  downloadBlob(new Blob([bytes], { type: 'image/png' }), `${safeFileName(title)}_${fileStamp()}.png`);
}
