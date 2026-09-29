/*
 * The SPFx build compiles to ES5, where a subclass of Error loses its
 * prototype: `instanceof` on these classes would always be false. Each class
 * therefore restores its prototype explicitly.
 */
function restorePrototype(instance: object, prototype: object): void {
  // Object.setPrototypeOf exists in every browser SharePoint supports; the ES5 typings just do not declare it.
  (Object as unknown as { setPrototypeOf(o: object, p: object): object }).setPrototypeOf(instance, prototype);
}

export class ScanCancelledError extends Error {
  constructor(message: string) {
    super(message);
    restorePrototype(this, ScanCancelledError.prototype);
    this.name = 'ScanCancelledError';
  }
}

export class HttpError extends Error {
  public status: number;

  constructor(status: number, message: string) {
    super(message);
    restorePrototype(this, HttpError.prototype);
    this.name = 'HttpError';
    this.status = status;
  }
}

/**
 * SharePoint has been throttling for longer than the scan is willing to wait
 * in one go. Not a failure: progress up to this point is kept, and the scan
 * can be resumed later.
 */
export class ScanPausedError extends Error {
  /** Progress so far (an IScanResult), attached by the scan so the caller can offer to resume. */
  public snapshot: unknown;

  constructor(message: string) {
    super(message);
    restorePrototype(this, ScanPausedError.prototype);
    this.name = 'ScanPausedError';
  }
}
