import { Route } from '../app/router';

/** The fixed pages, with the plain-language hint shown in the sidebar, switcher and help. */
export const PAGES: { route: Route; label: string; icon: string; hint: string }[] = [
  { route: { page: 'home' }, label: 'Home', icon: 'Home', hint: 'Your starting point: what needs attention and your recent boards' },
  { route: { page: 'mywork' }, label: 'My Work', icon: 'CheckList', hint: 'Everything assigned to you, on every board, by due date' },
  { route: { page: 'projects' }, label: 'My projects', icon: 'ProjectCollection', hint: 'Boards where you are project head, lead, sponsor or owner' },
  { route: { page: 'team' }, label: 'My team', icon: 'People', hint: 'What the people who report to you are working on' }
];

/** Simple ranking: starts-with beats contains; word starts beat the middle of a word. */
export function rankMatch(label: string, query: string): number {
  const l = label.toLowerCase();
  const q = query.trim().toLowerCase();
  if (!q) {
    return 1;
  }
  if (l.indexOf(q) === 0) {
    return 3;
  }
  if (l.split(/[\s\-_/]+/).some(w => w.indexOf(q) === 0)) {
    return 2;
  }
  return l.indexOf(q) >= 0 ? 1 : 0;
}


export interface ICrumb {
  label: string;
  route?: Route;
}

/** Where am I: Home › folder › board, or Home › My Work. The last crumb is the current page. */
export function crumbsFor(route: Route, boards: { id: number; title: string; folder: string }[]): ICrumb[] {
  const home: ICrumb = { label: 'Home', route: { page: 'home' } };
  if (route.page === 'home') {
    return [{ label: 'Home' }];
  }
  if (route.page === 'board') {
    const b = boards.filter(x => x.id === route.boardId)[0];
    const out: ICrumb[] = [home];
    if (b && b.folder) {
      out.push({ label: b.folder });
    }
    out.push(route.itemId ? { label: b ? b.title : 'Board', route: { page: 'board', boardId: route.boardId, view: route.view } } : { label: b ? b.title : 'Board' });
    if (route.itemId) {
      out.push({ label: 'Item' });
    }
    return out;
  }
  const page = PAGES.filter(p => p.route.page === route.page)[0];
  return [home, { label: page ? page.label : '' }];
}

