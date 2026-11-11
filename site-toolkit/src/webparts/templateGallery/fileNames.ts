/** SharePoint refuses these characters in file names. */
const INVALID_NAME = /["*:<>?/\\|]/;

export function validateFileName(name: string): string | undefined {
  const trimmed = name.trim();
  if (!trimmed) {
    return 'Give the new document a name.';
  }
  if (INVALID_NAME.test(trimmed)) {
    return 'A name cannot contain " * : < > ? / \\ |';
  }
  if (/^[.~]/.test(trimmed) || /\.$/.test(trimmed)) {
    return 'A name cannot start with a dot or tilde, or end with a dot.';
  }
  if (trimmed.length > 120) {
    return 'That name is too long.';
  }
  return undefined;
}

export function stem(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot > 0 ? fileName.substring(0, dot) : fileName;
}
