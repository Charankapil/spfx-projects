/**
 * Console logging that is safe for shared build logs.
 *
 * Scheduled runs often write to logs other people can read (the Actions log of a public
 * repository is public). Unless verbose logging is asked for, anything that looks like an
 * address or an e-mail is masked, so a log line can say what happened without saying where
 * or to whom.
 */

const URL_RE = /https?:\/\/[^\s"'<>)]+/gi;
const EMAIL_RE = /[A-Za-z0-9._%+#-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
const GUID_RE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;
const CLAIM_RE = /\b[ic]:0[^\s"'<>]*\|[^\s"'<>,)]+/gi;

export function redact(text: string): string {
  return text.replace(URL_RE, '<address>').replace(CLAIM_RE, '<login>').replace(EMAIL_RE, '<e-mail>').replace(GUID_RE, '<id>');
}

export interface Logger {
  verbose: boolean;
  info(msg: string): void;
  warn(msg: string): void;
  error(msg: string): void;
  /** Only shown with --verbose; never masked (the caller asked for detail). */
  debug(msg: string): void;
  /** Masks unless verbose. For text that goes somewhere other than the console. */
  safe(msg: string): string;
}

export function createLogger(verbose: boolean, sink: (line: string) => void = (l) => process.stderr.write(l + '\n')): Logger {
  const safe = (m: string): string => (verbose ? m : redact(m));
  const stamp = (): string => new Date().toISOString().substring(11, 19);
  return {
    verbose,
    info: (m) => sink(`${stamp()}  ${safe(m)}`),
    warn: (m) => sink(`${stamp()}  WARNING: ${safe(m)}`),
    error: (m) => sink(`${stamp()}  ERROR: ${safe(m)}`),
    debug: (m) => {
      if (verbose) {
        sink(`${stamp()}  . ${m}`);
      }
    },
    safe
  };
}

export const silentLogger: Logger = { verbose: false, info: () => undefined, warn: () => undefined, error: () => undefined, debug: () => undefined, safe: redact };
