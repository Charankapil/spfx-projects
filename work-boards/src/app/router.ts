import { BoardViewType } from '../models/types';

/**
 * Hash routes, so links to a board or item can be shared:
 *   #/                      home
 *   #/mywork                My Work
 *   #/projects              My projects (boards where I hold a project role)
 *   #/team                  My team (work of my reports)
 *   #/board/12              board 12, default view
 *   #/board/12/kanban       board 12, Kanban view
 *   #/board/12/table/item/5 board 12, table view, item 5 open
 */
export type Route =
  | { page: 'home' }
  | { page: 'mywork' }
  | { page: 'projects' }
  | { page: 'team' }
  | { page: 'board'; boardId: number; view?: BoardViewType; itemId?: number };

const VIEWS: BoardViewType[] = ['table', 'kanban', 'timeline'];

export function parseRoute(hash: string): Route {
  const parts = hash.replace(/^#\/?/, '').split('/').filter(p => p.length > 0);
  if (parts[0] === 'mywork') {
    return { page: 'mywork' };
  }
  if (parts[0] === 'projects') {
    return { page: 'projects' };
  }
  if (parts[0] === 'team') {
    return { page: 'team' };
  }
  if (parts[0] === 'board' && /^\d+$/.test(parts[1] || '')) {
    const route: Route = { page: 'board', boardId: parseInt(parts[1], 10) };
    if (parts[2] && VIEWS.indexOf(parts[2] as BoardViewType) >= 0) {
      route.view = parts[2] as BoardViewType;
    }
    if (parts[3] === 'item' && /^\d+$/.test(parts[4] || '')) {
      route.itemId = parseInt(parts[4], 10);
    }
    return route;
  }
  return { page: 'home' };
}

export function routeToHash(route: Route): string {
  if (route.page === 'mywork') {
    return '#/mywork';
  }
  if (route.page === 'projects') {
    return '#/projects';
  }
  if (route.page === 'team') {
    return '#/team';
  }
  if (route.page === 'board') {
    let h = '#/board/' + route.boardId;
    if (route.view || route.itemId) {
      h += '/' + (route.view || 'table');
    }
    if (route.itemId) {
      h += '/item/' + route.itemId;
    }
    return h;
  }
  return '#/';
}
