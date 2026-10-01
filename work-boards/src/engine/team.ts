import { IPersonWorkRow, IPerson } from '../models/types';

export interface IPersonTotals {
  person: IPerson;
  open: number;
  overdue: number;
  dueThisWeek: number;
  done: number;
  boards: number;
}

/**
 * Per-person totals for My team. An item with two team members counts for each of them.
 * `today` and `weekEnd` are "YYYY-MM-DD".
 */
export function teamTotals(rows: IPersonWorkRow[], people: IPerson[], today: string, weekEnd: string): IPersonTotals[] {
  return people.map(person => {
    const mine = rows.filter(r => r.people.some(p => p.id === person.id));
    const boardIds: number[] = [];
    let open = 0;
    let overdue = 0;
    let dueThisWeek = 0;
    let done = 0;
    mine.forEach(r => {
      if (boardIds.indexOf(r.board.id) < 0) {
        boardIds.push(r.board.id);
      }
      if (r.status && r.status.isDone) {
        done++;
        return;
      }
      open++;
      if (r.due && r.due < today) {
        overdue++;
      } else if (r.due && r.due <= weekEnd) {
        dueThisWeek++;
      }
    });
    return { person, open, overdue, dueThisWeek, done, boards: boardIds.length };
  });
}

/** Open items first, overdue first, then by due date (no date last), then title. */
export function sortTeamRows(rows: IPersonWorkRow[]): IPersonWorkRow[] {
  return rows.slice().sort((a, b) => {
    const ad = a.status && a.status.isDone ? 1 : 0;
    const bd = b.status && b.status.isDone ? 1 : 0;
    if (ad !== bd) {
      return ad - bd;
    }
    return (a.due || '9999').localeCompare(b.due || '9999') || a.item.title.localeCompare(b.item.title);
  });
}
