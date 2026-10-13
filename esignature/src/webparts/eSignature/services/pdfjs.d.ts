// Minimal typings for the parts of pdf.js the web part uses. Declared here so
// the build does not depend on pdf.js's own .d.ts files, which need a newer
// TypeScript than SPFx 1.18 ships with.
declare module 'pdfjs-dist/legacy/build/pdf' {
  export interface IPdfViewport {
    width: number;
    height: number;
  }
  export interface IPdfPage {
    getViewport(options: { scale: number }): IPdfViewport;
    render(options: { canvasContext: CanvasRenderingContext2D; viewport: IPdfViewport }): { promise: Promise<void> };
    cleanup(): void;
  }
  export interface IPdfDocument {
    numPages: number;
    getPage(pageNumber: number): Promise<IPdfPage>;
    destroy(): Promise<void>;
  }
  export function getDocument(source: { data: Uint8Array; isEvalSupported?: boolean }): { promise: Promise<IPdfDocument> };
  export const GlobalWorkerOptions: { workerSrc: string };
}

declare module 'pdfjs-dist/legacy/build/pdf.worker.entry';
