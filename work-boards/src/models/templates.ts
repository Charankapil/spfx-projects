import { IBoardColumn, IBoardGroup, ILabel, CellValue, BoardViewType } from './types';
import { F } from '../engine/fieldMap';

export interface ITemplateItem {
  title: string;
  group: string;
  /** Values by template column id. Dates as day offsets from today: "+3", "-2". */
  values: { [columnId: string]: CellValue | string };
  subitems?: string[];
}

export interface IBoardTemplate {
  id: string;
  title: string;
  description: string;
  icon: string;
  /** Columns with a template-local id. `field` set means a fixed field; otherwise a new field is created. */
  columns: (Omit<IBoardColumn, 'field'> & { field?: string })[];
  groups: IBoardGroup[];
  kanbanColumnId?: string;
  timelineColumnId?: string;
  defaultView: BoardViewType;
  items: ITemplateItem[];
}

export const STATUS_LABELS: ILabel[] = [
  { id: 'l_ns', text: 'Not started', color: '#c4c4c4' },
  { id: 'l_wo', text: 'Working on it', color: '#fdab3d' },
  { id: 'l_st', text: 'Stuck', color: '#e2445c' },
  { id: 'l_dn', text: 'Done', color: '#00c875', isDone: true }
];

export const PRIORITY_LABELS: ILabel[] = [
  { id: 'p_cr', text: 'Critical', color: '#333333' },
  { id: 'p_hi', text: 'High', color: '#401694' },
  { id: 'p_md', text: 'Medium', color: '#5559df' },
  { id: 'p_lo', text: 'Low', color: '#579bfc' }
];

const owner = { id: 'owner', title: 'Owner', type: 'people' as const, field: F.Owner };
const status = { id: 'status', title: 'Status', type: 'status' as const, field: F.Status, labels: STATUS_LABELS };
const priority = { id: 'priority', title: 'Priority', type: 'status' as const, labels: PRIORITY_LABELS };
const due = { id: 'due', title: 'Due date', type: 'date' as const, field: F.DueDate };
const timeline = { id: 'timeline', title: 'Timeline', type: 'timeline' as const, field: F.StartDate, fieldEnd: F.DueDate };

export const TEMPLATES: IBoardTemplate[] = [
  {
    id: 'project',
    title: 'Project plan',
    description: 'Phases, owners, status, priority and a timeline you can see as a Gantt chart.',
    icon: 'Timeline',
    columns: [
      owner,
      status,
      priority,
      timeline,
      { id: 'effort', title: 'Effort', type: 'number', unit: 'h', decimals: 0 },
      { id: 'notes', title: 'Notes', type: 'text' }
    ],
    groups: [
      { id: 'g_plan', title: 'Planning', color: '#579bfc' },
      { id: 'g_exec', title: 'Execution', color: '#a25ddc' },
      { id: 'g_close', title: 'Closing', color: '#00c875' }
    ],
    kanbanColumnId: 'status',
    timelineColumnId: 'timeline',
    defaultView: 'table',
    items: [
      { title: 'Agree scope and goals', group: 'g_plan', values: { status: 'Done', priority: 'High', timeline: { start: '-7', end: '-3' }, effort: 6 } },
      { title: 'Build the project schedule', group: 'g_plan', values: { status: 'Working on it', priority: 'Medium', timeline: { start: '-2', end: '+2' }, effort: 4 }, subitems: ['List milestones', 'Confirm dates with the team'] },
      { title: 'Deliver phase 1', group: 'g_exec', values: { status: 'Not started', priority: 'High', timeline: { start: '+3', end: '+17' }, effort: 40 } },
      { title: 'Weekly status review', group: 'g_exec', values: { status: 'Not started', priority: 'Low', timeline: { start: '+7', end: '+7' } } },
      { title: 'Lessons learned', group: 'g_close', values: { status: 'Not started', priority: 'Low', timeline: { start: '+21', end: '+22' } } }
    ]
  },
  {
    id: 'tasks',
    title: 'Team tasks',
    description: 'A simple to-do board for the team, grouped by when work is due.',
    icon: 'TaskManager',
    columns: [owner, status, priority, due],
    groups: [
      { id: 'g_week', title: 'This week', color: '#579bfc' },
      { id: 'g_next', title: 'Next week', color: '#a25ddc' },
      { id: 'g_later', title: 'Later', color: '#fdab3d' }
    ],
    kanbanColumnId: 'status',
    timelineColumnId: 'due',
    defaultView: 'table',
    items: [
      { title: 'Prepare the monthly report', group: 'g_week', values: { status: 'Working on it', priority: 'High', due: '+2' } },
      { title: 'Update the team page', group: 'g_week', values: { status: 'Not started', priority: 'Medium', due: '+4' } },
      { title: 'Plan the team day', group: 'g_next', values: { status: 'Not started', priority: 'Low', due: '+9' } }
    ]
  },
  {
    id: 'requests',
    title: 'Team requests',
    description: 'Track incoming requests from intake to done, with who asked and what type.',
    icon: 'Inbox',
    columns: [
      { id: 'requester', title: 'Requested by', type: 'people' },
      owner,
      {
        id: 'status', title: 'Status', type: 'status', field: F.Status,
        labels: [
          { id: 'r_new', text: 'New', color: '#579bfc' },
          { id: 'r_rev', text: 'In review', color: '#a25ddc' },
          { id: 'r_wo', text: 'In progress', color: '#fdab3d' },
          { id: 'r_wait', text: 'Waiting', color: '#e2445c' },
          { id: 'r_dn', text: 'Done', color: '#00c875', isDone: true }
        ]
      },
      {
        id: 'type', title: 'Request type', type: 'dropdown',
        labels: [
          { id: 't_q', text: 'Question', color: '#579bfc' },
          { id: 't_c', text: 'Change', color: '#fdab3d' },
          { id: 't_n', text: 'New work', color: '#00c875' }
        ]
      },
      due,
      { id: 'details', title: 'Details', type: 'longtext' }
    ],
    groups: [
      { id: 'g_new', title: 'New requests', color: '#579bfc' },
      { id: 'g_prog', title: 'In progress', color: '#fdab3d' },
      { id: 'g_done', title: 'Completed', color: '#00c875' }
    ],
    kanbanColumnId: 'status',
    timelineColumnId: 'due',
    defaultView: 'kanban',
    items: [
      { title: 'Access to the shared drive', group: 'g_new', values: { status: 'New', type: ['Question'], due: '+3' } },
      { title: 'Change the report layout', group: 'g_prog', values: { status: 'In progress', type: ['Change'], due: '+5' } }
    ]
  },
  {
    id: 'blank',
    title: 'Blank board',
    description: 'Start with Owner, Status and Due date, and add your own columns.',
    icon: 'Add',
    columns: [owner, status, due],
    groups: [
      { id: 'g_todo', title: 'To do', color: '#579bfc' },
      { id: 'g_done', title: 'Done', color: '#00c875' }
    ],
    kanbanColumnId: 'status',
    timelineColumnId: 'due',
    defaultView: 'table',
    items: []
  }
];

export function getTemplate(id: string): IBoardTemplate {
  return TEMPLATES.filter(t => t.id === id)[0] || TEMPLATES[TEMPLATES.length - 1];
}
