export interface ISearchCell { Key: string; Value: unknown; }
export interface ISearchRow { Cells?: ISearchCell[]; }

/** Flattens SharePoint Search result rows ({ Cells: [{ Key, Value }] }) into plain property bags. */
export function mapSearchRows(rows: ISearchRow[] | undefined): Record<string, string>[] {
  return (rows || []).map((row: ISearchRow) => {
    const cells: Record<string, string> = {};
    (row.Cells || []).forEach((cell: ISearchCell) => {
      cells[cell.Key] = cell.Value === null || cell.Value === undefined ? '' : String(cell.Value);
    });
    return cells;
  });
}
