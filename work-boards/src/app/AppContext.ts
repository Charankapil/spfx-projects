import * as React from 'react';
import { SpClient } from '../services/SpClient';
import { BoardService } from '../services/BoardService';
import { ItemService } from '../services/ItemService';
import { UpdateService } from '../services/UpdateService';
import { PeopleService } from '../services/PeopleService';
import { PrefsService } from '../services/PrefsService';
import { IBoard, IPerson, IUserPrefs } from '../models/types';
import { Route } from './router';

export interface IServices {
  sp: SpClient;
  boards: BoardService;
  items: ItemService;
  updates: UpdateService;
  people: PeopleService;
  prefs: PrefsService;
}

export interface IAppContext {
  services: IServices;
  me: IPerson;
  siteTitle: string;
  webUrl: string;
  /** The user can manage the site (set up, upgrade, create private boards). */
  isSiteOwner: boolean;
  boards: IBoard[];
  prefs: IUserPrefs;
  reloadBoards: () => Promise<IBoard[]>;
  /** Replace one board in the shared list after a change. */
  replaceBoard: (board: IBoard) => void;
  toggleFavourite: (boardId: number) => void;
  navigate: (route: Route) => void;
}

export const AppContext = React.createContext<IAppContext>(undefined as unknown as IAppContext);

export function useApp(): IAppContext {
  return React.useContext(AppContext);
}
