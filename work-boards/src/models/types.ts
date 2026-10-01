/** Column types a board can hold. Each maps to one (or two) real SharePoint fields. */
export type ColumnType =
  | 'status'
  | 'people'
  | 'text'
  | 'longtext'
  | 'number'
  | 'date'
  | 'timeline'
  | 'dropdown'
  | 'checkbox'
  | 'link';

/** A coloured label used by Status and Dropdown columns. */
export interface ILabel {
  id: string;
  text: string;
  color: string;
  /** Status only: marks the label that means "finished" (used by My Work, progress and overdue). */
  isDone?: boolean;
}

export interface IBoardColumn {
  id: string;
  title: string;
  type: ColumnType;
  /** Internal name of the SharePoint field. Timeline uses this for the start date. */
  field: string;
  /** Timeline only: internal name of the end date field. */
  fieldEnd?: string;
  width?: number;
  hidden?: boolean;
  labels?: ILabel[];
  /** Numbers only. */
  unit?: string;
  decimals?: number;
}

export interface IBoardGroup {
  id: string;
  title: string;
  color: string;
}

export type BoardViewType = 'table' | 'kanban' | 'timeline';

/** Stored as JSON in the board registry row. */
export interface IBoardConfig {
  v: 1;
  columns: IBoardColumn[];
  groups: IBoardGroup[];
  /** Column id Kanban lanes come from (a Status column). */
  kanbanColumnId?: string;
  /** Column id the Timeline view reads (a Timeline or Date column). */
  timelineColumnId?: string;
  defaultView?: BoardViewType;
  /** True once the board list has the WB_Links field (boards created from v0.1.5 on, or upgraded by an owner). */
  linksField?: boolean;
}

export type BoardPrivacy = 'Main' | 'Private';

export interface IBoard {
  /** Item id in the WB_Boards registry list. */
  id: number;
  title: string;
  key: string;
  description: string;
  folder: string;
  color: string;
  privacy: BoardPrivacy;
  itemsListId: string;
  updatesListId: string;
  ownerIds: number[];
  /** Optional project roles, set in board settings. */
  roles: IBoardRoles;
  archived: boolean;
  config: IBoardConfig;
  etag: string;
}

export type BoardRole = 'head' | 'lead' | 'sponsor';

/** Site user ids per project role. */
export interface IBoardRoles {
  head: number[];
  lead: number[];
  sponsor: number[];
}

export const ROLE_LABELS: { [role in BoardRole]: string } = {
  head: 'Project head',
  lead: 'Project lead',
  sponsor: 'Project sponsor'
};

/** A file linked to an item: in a SharePoint library, OneDrive or anywhere else. */
export interface ILink {
  name: string;
  url: string;
  addedBy?: string;
  added?: string;
}

export interface IPerson {
  id: number;
  title: string;
  email: string;
}

export interface ITimelineValue {
  start: string | null;
  end: string | null;
}

/** Cell values, keyed by column id. Dates are "YYYY-MM-DD" strings. */
export type CellValue = string | number | boolean | string[] | IPerson[] | ITimelineValue | null;

export interface IWorkItem {
  id: number;
  title: string;
  groupId: string;
  sortOrder: number;
  parentId: number | null;
  values: { [columnId: string]: CellValue };
  created: string;
  modified: string;
  author: IPerson | null;
  editor: IPerson | null;
  attachments: boolean;
  links: ILink[];
  etag: string;
}

export interface IUpdate {
  id: number;
  itemId: number;
  parentId: number | null;
  body: string;
  author: IPerson | null;
  created: string;
  mentionIds: number[];
  likedByIds: number[];
  etag: string;
}

export interface IActivityEntry {
  version: string;
  when: string;
  who: string;
  /** True for the first version: the item was created. */
  created: boolean;
  changes: { column: string; from: string; to: string }[];
}

export interface IAttachment {
  fileName: string;
  url: string;
}

/** Row in My Work: an item plus the board it lives on. */
export interface IMyWorkRow {
  board: IBoard;
  item: IWorkItem;
  status: ILabel | null;
  due: string | null;
}

export interface IUserPrefs {
  favourites: number[];
  recent: number[];
  /** People added by hand to My team (site user ids). */
  team: number[];
}

/** An item that involves a person, for My team. */
export interface IPersonWorkRow {
  board: IBoard;
  item: IWorkItem;
  /** People from the watched set who appear in any People column of the item. */
  people: IPerson[];
  status: ILabel | null;
  due: string | null;
}
