import { keyBetween, keyForIndex, needsRebalance, rebalance, ORDER_STEP } from '../ordering';
import { toSpDate, fromSpDate, addDays, diffDays, formatDay, formatRange, relativeDay, todayIso } from '../dates';
import { fieldXml, readItem, writeCell, cellText, itemQueryParts, xmlEscape, F, peopleIdsIn, readIds } from '../fieldMap';
import { diffVersions, versionField } from '../activity';
import { parseSegments, mentionIds, mentionToken, plainText } from '../mentions';
import { boardToCsv, csvEscape } from '../csv';
import { visibleRows, EMPTY_VIEW, summarize, dueOf, isDone, activeFilterCount } from '../viewQuery';
import { suggestBoardKey, isValidBoardKey, shortId } from '../ids';
import { parseRoute, routeToHash } from '../../app/router';
import { hasPermission, PermissionKind, boardRights } from '../../services/permissions';
import { pickFieldsForNewColumn, parseConfig } from '../../services/BoardService';
import { IBoard, IBoardColumn, IWorkItem, IBoardConfig } from '../../models/types';
import { STATUS_LABELS } from '../../models/templates';

const statusCol: IBoardColumn = { id: 'status', title: 'Status', type: 'status', field: F.Status, labels: STATUS_LABELS };
const ownerCol: IBoardColumn = { id: 'owner', title: 'Owner', type: 'people', field: F.Owner };
const timelineCol: IBoardColumn = { id: 'tl', title: 'Timeline', type: 'timeline', field: F.StartDate, fieldEnd: F.DueDate };
const tagsCol: IBoardColumn = { id: 'tags', title: 'Tags', type: 'dropdown', field: 'WB_c_tags01', labels: [{ id: 'a', text: 'Urgent', color: '#e2445c' }] };
const effortCol: IBoardColumn = { id: 'effort', title: 'Effort', type: 'number', field: 'WB_c_eff001', unit: 'h' };

const config: IBoardConfig = {
  v: 1,
  columns: [ownerCol, statusCol, timelineCol, tagsCol, effortCol],
  groups: [{ id: 'g1', title: 'Planning', color: '#579bfc' }, { id: 'g2', title: 'Done', color: '#00c875' }],
  kanbanColumnId: 'status',
  timelineColumnId: 'tl'
};

const board: IBoard = {
  id: 1, title: 'Launch', key: 'LCH', description: '', folder: '', color: '#579bfc', privacy: 'Main',
  itemsListId: 'x', updatesListId: 'y', ownerIds: [7], archived: false, config, etag: '"1"'
};

function item(id: number, title: string, groupId: string, sortOrder: number, values: IWorkItem['values'], parentId: number | null = null): IWorkItem {
  return { id, title, groupId, sortOrder, parentId, values, created: '2026-09-01T10:00:00Z', modified: '2026-09-01T10:00:00Z', author: null, editor: null, attachments: false, etag: '"1"' };
}

describe('ordering', () => {
  it('puts keys between neighbours', () => {
    expect(keyBetween(null, null)).toBe(ORDER_STEP);
    expect(keyBetween(null, 1024)).toBe(0);
    expect(keyBetween(1024, null)).toBe(2048);
    expect(keyBetween(1024, 2048)).toBe(1536);
  });

  it('computes the key for an index and flags tiny gaps', () => {
    const sibs = [{ sortOrder: 1024 }, { sortOrder: 2048 }];
    expect(keyForIndex(sibs, 0).key).toBe(0);
    expect(keyForIndex(sibs, 1).key).toBe(1536);
    expect(keyForIndex(sibs, 2).key).toBe(3072);
    expect(needsRebalance(1, 1 + 1e-9)).toBe(true);
    expect(needsRebalance(1, 2)).toBe(false);
  });

  it('rebalances only changed keys', () => {
    expect(rebalance([{ id: 1, sortOrder: 1024 }, { id: 2, sortOrder: 5 }])).toEqual([{ id: 2, sortOrder: 2048 }]);
  });
});

describe('dates', () => {
  it('writes dates at noon UTC and reads the calendar day back', () => {
    expect(toSpDate('2026-10-03')).toBe('2026-10-03T12:00:00Z');
    expect(toSpDate(null)).toBeNull();
    expect(fromSpDate('2026-10-03T12:00:00Z')).toBe('2026-10-03');
    expect(fromSpDate(null)).toBeNull();
  });

  it('does day arithmetic across month ends', () => {
    expect(addDays('2026-01-30', 3)).toBe('2026-02-02');
    expect(addDays('2026-03-01', -1)).toBe('2026-02-28');
    expect(diffDays('2026-09-28', '2026-10-05')).toBe(7);
  });

  it('formats days and ranges', () => {
    const now = new Date(2026, 8, 29);
    expect(formatDay('2026-10-03', now)).toBe('3 Oct');
    expect(formatDay('2027-01-03', now)).toBe('3 Jan 2027');
    expect(formatRange('2026-10-01', '2026-10-08', now)).toBe('1 Oct – 8 Oct');
    expect(formatRange('2026-10-01', '2026-10-01', now)).toBe('1 Oct');
    expect(relativeDay('2026-09-29', now)).toBe('Today');
    expect(relativeDay('2026-09-30', now)).toBe('Tomorrow');
    expect(relativeDay('2026-09-28', now)).toBe('Yesterday');
    expect(todayIso(now)).toBe('2026-09-29');
  });
});

describe('fieldMap', () => {
  it('builds field XML with escaping, indexes and fill-in choices', () => {
    const xml = fieldXml('status', 'WB_Status', 'Status & "state"', { indexed: true, choices: ['Done', 'A<B'] });
    expect(xml).toContain('Type="Choice"');
    expect(xml).toContain('FillInChoice="TRUE"');
    expect(xml).toContain('Indexed="TRUE"');
    expect(xml).toContain('DisplayName="Status &amp; &quot;state&quot;"');
    expect(xml).toContain('<CHOICE>A&lt;B</CHOICE>');
    expect(fieldXml('people', 'WB_Owner', 'Owner')).toContain('Type="UserMulti"');
    expect(fieldXml('date', 'WB_DueDate', 'Due', { indexed: true })).toContain('Format="DateOnly"');
    expect(xmlEscape("it's")).toBe('it&apos;s');
  });

  it('selects and expands the right fields', () => {
    const q = itemQueryParts(config.columns);
    // People are read as ids, never projected: SharePoint rejects WB_Owner/EMail.
    expect(q.expand).not.toContain(F.Owner);
    expect(q.select).toContain(`${F.Owner}Id`);
    expect(q.select.some(s => s.indexOf(F.Owner + '/') === 0)).toBe(false);
    expect(q.select).toContain(F.StartDate);
    expect(q.select).toContain(F.DueDate);
    expect(q.select).toContain('WB_c_tags01');
  });

  it('reads a REST row into an item', () => {
    const row = {
      Id: 5, Title: 'Write brief', [F.GroupId]: 'g1', [F.SortOrder]: 2048, [F.ParentId]: null,
      [F.Owner + 'Id']: [7, 9],
      [F.Status]: 'Working on it', [F.StartDate]: '2026-10-01T12:00:00Z', [F.DueDate]: '2026-10-08T12:00:00Z',
      WB_c_tags01: ['Urgent'], WB_c_eff001: 6, Created: '2026-09-01T10:00:00Z', Modified: '2026-09-02T10:00:00Z',
      Attachments: true, 'odata.etag': '"3"'
    };
    const users = new Map([[7, { id: 7, title: 'Anna Smith', email: 'anna@contoso.com' }]]);
    const it1 = readItem(config.columns, row, users);
    expect(it1.id).toBe(5);
    expect(it1.parentId).toBeNull();
    expect(it1.etag).toBe('"3"');
    expect(it1.values.owner).toEqual([{ id: 7, title: 'Anna Smith', email: 'anna@contoso.com' }, { id: 9, title: 'User 9', email: '' }]);
    expect(peopleIdsIn(config.columns, [row, { [F.Owner + 'Id']: { results: [9, 11] } }])).toEqual([7, 9, 11]);
    expect(readIds(null)).toEqual([]);
    expect(readIds(5)).toEqual([5]);
    expect(it1.values.tl).toEqual({ start: '2026-10-01', end: '2026-10-08' });
    expect(it1.values.tags).toEqual(['Urgent']);
    expect(it1.values.effort).toBe(6);
  });

  it('writes verbose OData payloads', () => {
    expect(writeCell(ownerCol, [{ id: 3, title: 'A', email: '' }])).toEqual({ WB_OwnerId: { results: [3] } });
    expect(writeCell(tagsCol, ['Urgent'])).toEqual({ WB_c_tags01: { __metadata: { type: 'Collection(Edm.String)' }, results: ['Urgent'] } });
    expect(writeCell(timelineCol, { start: '2026-10-01', end: null })).toEqual({ WB_StartDate: '2026-10-01T12:00:00Z', WB_DueDate: null });
    expect(writeCell(effortCol, '')).toEqual({ WB_c_eff001: null });
  });

  it('turns cells into text', () => {
    expect(cellText(ownerCol, [{ id: 1, title: 'A', email: '' }, { id: 2, title: 'B', email: '' }])).toBe('A, B');
    expect(cellText(timelineCol, { start: null, end: null })).toBe('');
  });
});

describe('activity', () => {
  it('reads encoded field names from the versions endpoint', () => {
    expect(versionField({ VersionLabel: '1.0', Created: '', WB_x005f_Status: 'Done' }, 'WB_Status')).toBe('Done');
  });

  it('diffs versions into readable changes, newest first', () => {
    const versions = [
      { VersionLabel: '2.0', Created: '2026-09-02T10:00:00Z', Editor: { LookupValue: 'Ben' }, Title: 'Brief', WB_x005f_GroupId: 'g2', WB_x005f_Status: 'Done', WB_x005f_Owner: [{ LookupValue: 'Anna' }] },
      { VersionLabel: '1.0', Created: '2026-09-01T10:00:00Z', Editor: { LookupValue: 'Anna' }, Title: 'Brief', WB_x005f_GroupId: 'g1', WB_x005f_Status: 'Working on it', WB_x005f_Owner: [] }
    ];
    const entries = diffVersions(config, versions);
    expect(entries).toHaveLength(2);
    expect(entries[1].created).toBe(true);
    expect(entries[0].who).toBe('Ben');
    expect(entries[0].changes).toEqual(expect.arrayContaining([
      { column: 'Group', from: 'Planning', to: 'Done' },
      { column: 'Status', from: 'Working on it', to: 'Done' },
      { column: 'Owner', from: '', to: 'Anna' }
    ]));
  });
});

describe('mentions', () => {
  it('parses and builds mention tokens', () => {
    const body = `Hi ${mentionToken('Anna Smith', 12)}, can you check? cc ${mentionToken('Ben', 4)}`;
    expect(mentionIds(body)).toEqual([12, 4]);
    expect(plainText(body)).toBe('Hi @Anna Smith, can you check? cc @Ben');
    expect(parseSegments(body)[1]).toEqual({ kind: 'mention', id: 12, name: 'Anna Smith' });
    expect(parseSegments('no mentions')).toEqual([{ kind: 'text', text: 'no mentions' }]);
  });
});

describe('csv', () => {
  it('escapes values and writes items in board order with subitems', () => {
    expect(csvEscape('a,b')).toBe('"a,b"');
    expect(csvEscape('say "hi"')).toBe('"say ""hi"""');
    const items = [
      item(2, 'Second', 'g1', 2048, { status: 'Done' }),
      item(1, 'First', 'g1', 1024, { status: 'Stuck' }),
      item(3, 'Sub of first', 'g1', 1024, {}, 1)
    ];
    const lines = boardToCsv(board, items).replace('﻿', '').split('\r\n');
    expect(lines[0]).toBe('Group,Item,Parent item,ID,Owner,Status,Timeline,Tags,Effort');
    expect(lines[1]).toContain('Planning,First,,LCH-1');
    expect(lines[2]).toContain('Planning,Sub of first,First,LCH-3');
    expect(lines[3]).toContain('Planning,Second,,LCH-2');
  });
});

describe('viewQuery', () => {
  const now = new Date(2026, 8, 29);
  const items = [
    item(1, 'Write brief', 'g1', 1024, { status: 'Done', owner: [{ id: 7, title: 'Anna', email: '' }], tl: { start: '2026-09-20', end: '2026-09-25' }, effort: 3 }),
    item(2, 'Build page', 'g1', 2048, { status: 'Working on it', owner: [], tl: { start: '2026-09-20', end: '2026-09-27' }, effort: 5 }),
    item(3, 'Launch', 'g2', 1024, { status: null, tl: { start: '2026-10-02', end: '2026-10-02' } }),
    item(4, 'Fix typo', 'g1', 1024, { status: 'Stuck' }, 2),
    item(5, 'Orphan group item', 'gone', 4096, {})
  ];

  it('groups items, nests subitems and puts unknown groups in the first group', () => {
    const rows = visibleRows(board, items, EMPTY_VIEW, now);
    expect(rows.g1.map(r => r.item.id)).toEqual([1, 2, 5]);
    expect(rows.g1[1].subitems.map(s => s.id)).toEqual([4]);
    expect(rows.g2.map(r => r.item.id)).toEqual([3]);
  });

  it('filters by search, person, labels and dates', () => {
    expect(visibleRows(board, items, { ...EMPTY_VIEW, search: 'brief' }, now).g1.map(r => r.item.id)).toEqual([1]);
    // A parent shows when a subitem matches.
    expect(visibleRows(board, items, { ...EMPTY_VIEW, search: 'typo' }, now).g1.map(r => r.item.id)).toEqual([2]);
    expect(visibleRows(board, items, { ...EMPTY_VIEW, personId: 7 }, now).g1.map(r => r.item.id)).toEqual([1]);
    const stuckOrEmpty = visibleRows(board, items, { ...EMPTY_VIEW, labelFilters: { status: ['Stuck', ''] } }, now);
    expect(stuckOrEmpty.g2.map(r => r.item.id)).toEqual([3]);
    expect(stuckOrEmpty.g1.map(r => r.item.id)).toEqual([2, 5]);
    // Overdue excludes done items.
    expect(visibleRows(board, items, { ...EMPTY_VIEW, dateFilter: 'overdue' }, now).g1.map(r => r.item.id)).toEqual([2]);
    expect(visibleRows(board, items, { ...EMPTY_VIEW, dateFilter: 'thisweek' }, now).g2.map(r => r.item.id)).toEqual([3]);
    expect(activeFilterCount({ ...EMPTY_VIEW, dateFilter: 'overdue', labelFilters: { status: ['Done'], x: [] } })).toBe(2);
  });

  it('sorts by a column', () => {
    const rows = visibleRows(board, items, { ...EMPTY_VIEW, sort: { columnId: 'effort', dir: 'desc' } }, now);
    expect(rows.g1.map(r => r.item.id)).toEqual([2, 1, 5]);
  });

  it('summarises columns and reads due and done', () => {
    const s = summarize(statusCol, items.slice(0, 3));
    expect(s.kind === 'labels' && s.parts.map(p => p.label.text)).toEqual(['Working on it', 'Done', 'Empty']);
    const sum = summarize(effortCol, items);
    expect(sum.kind === 'sum' && sum.value).toBe(8);
    const range = summarize(timelineCol, items);
    expect(range).toEqual({ kind: 'range', start: '2026-09-20', end: '2026-10-02' });
    expect(dueOf(board, items[0])).toBe('2026-09-25');
    expect(isDone(board, items[0])).toBe(true);
    expect(isDone(board, items[1])).toBe(false);
  });
});

describe('ids and routes', () => {
  it('suggests unique, valid board keys', () => {
    expect(suggestBoardKey('Q4 Marketing Campaign', [])).toBe('QMC');
    expect(suggestBoardKey('Q4 Marketing Campaign', ['QMC'])).toBe('QMC2');
    expect(suggestBoardKey('Onboarding', [])).toBe('ONBO');
    expect(isValidBoardKey('QMC2')).toBe(true);
    expect(isValidBoardKey('2QMC')).toBe(false);
    expect(/^[a-z][a-z0-9]{5}$/.test(shortId())).toBe(true);
  });

  it('round-trips routes', () => {
    expect(parseRoute('#/board/12/kanban/item/5')).toEqual({ page: 'board', boardId: 12, view: 'kanban', itemId: 5 });
    expect(parseRoute('#/mywork')).toEqual({ page: 'mywork' });
    expect(parseRoute('#/nonsense')).toEqual({ page: 'home' });
    expect(routeToHash({ page: 'board', boardId: 12, itemId: 5 })).toBe('#/board/12/table/item/5');
    expect(routeToHash({ page: 'board', boardId: 12 })).toBe('#/board/12');
  });
});

describe('permissions', () => {
  it('reads permission bits from High/Low', () => {
    // Edit level: view, add, edit, delete and manage lists.
    const edit = { High: '432', Low: String(1 + 2 + 4 + 8 + 2048) };
    expect(hasPermission(edit, PermissionKind.EditListItems)).toBe(true);
    expect(hasPermission(edit, PermissionKind.ManageLists)).toBe(true);
    expect(hasPermission(edit, PermissionKind.ManagePermissions)).toBe(false);
    const full = { High: '2147483647', Low: '4294967295' };
    expect(hasPermission(full, PermissionKind.ManageWeb)).toBe(true);
    const read = { High: '0', Low: '1' };
    expect(boardRights(read, true)).toEqual(expect.objectContaining({ canView: true, canEdit: false, canManage: false }));
  });
});

describe('board config', () => {
  it('reuses free fixed fields for new columns', () => {
    const empty: IBoardConfig = { v: 1, columns: [], groups: [] };
    expect(pickFieldsForNewColumn(empty, 'people')).toEqual({ field: F.Owner });
    expect(pickFieldsForNewColumn(empty, 'timeline')).toEqual({ field: F.StartDate, fieldEnd: F.DueDate });
    expect(pickFieldsForNewColumn(config, 'people').field).toMatch(/^WB_c_[a-z0-9]{6}$/);
    const tl = pickFieldsForNewColumn(config, 'timeline');
    expect(tl.fieldEnd).toBe(tl.field + 'e');
  });

  it('parses stored config with safe defaults', () => {
    const cfg = parseConfig('not json');
    expect(cfg.groups).toHaveLength(1);
    expect(cfg.columns).toEqual([]);
    expect(parseConfig(JSON.stringify(config)).kanbanColumnId).toBe('status');
  });
});
