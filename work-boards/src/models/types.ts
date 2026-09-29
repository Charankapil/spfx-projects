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
  archived: boolean;
  config: IBoardConfig;
  etag: string;
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
}
