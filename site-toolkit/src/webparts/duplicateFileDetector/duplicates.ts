import { fileNameOf } from '../../common/search';

export interface ICopy {
  path: string;
  modified: string;
  modifiedBy: string;
}

export interface IDuplicateGroup {
  name: string;
  size: number;
  copies: ICopy[];
}

/**
 * Groups files that share a name and a size. Files of zero bytes are
 * ignored: they are all "the same" without being copies of anything.
 */
export function collectDuplicates(groups: Record<string, IDuplicateGroup>): IDuplicateGroup[] {
  const out: IDuplicateGroup[] = [];
  for (const key of Object.keys(groups)) {
    if (groups[key].copies.length > 1) {
      out.push(groups[key]);
    }
  }
  return out.sort((a, b) => b.size * (b.copies.length - 1) - a.size * (a.copies.length - 1));
}

export function addToGroups(groups: Record<string, IDuplicateGroup>, path: string, size: number, modified: string, modifiedBy: string): void {
  if (!path || size <= 0) {
    return;
  }
  const name = fileNameOf(path);
  const key = `${name.toLowerCase()}|${size}`;
  if (!groups[key]) {
    groups[key] = { name, size, copies: [] };
  }
  // The same document can be returned twice by the index; one path counts once.
  if (!groups[key].copies.some((copy) => copy.path === path)) {
    groups[key].copies.push({ path, modified, modifiedBy });
  }
}
