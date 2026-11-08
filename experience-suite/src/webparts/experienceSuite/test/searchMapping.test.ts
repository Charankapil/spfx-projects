import { mapSearchRows } from '../utils/searchMapping';

describe('mapSearchRows', () => {
  it('maps SharePoint search cells into a property object', () => {
    expect(mapSearchRows([{ Cells: [{ Key: 'Title', Value: 'Policy' }, { Key: 'Path', Value: 'https://tenant/policy' }] }])).toEqual([
      { Title: 'Policy', Path: 'https://tenant/policy' }
    ]);
  });

  it('normalizes null/undefined values and empty rows', () => {
    expect(mapSearchRows([{ Cells: [{ Key: 'Department', Value: null }, { Key: 'Skills', Value: undefined }] }, {}])).toEqual([{ Department: '', Skills: '' }, {}]);
  });

  it('stringifies non-string values', () => {
    expect(mapSearchRows([{ Cells: [{ Key: 'Rank', Value: 12.5 }] }])).toEqual([{ Rank: '12.5' }]);
  });

  it('returns an empty array for missing results', () => {
    expect(mapSearchRows([])).toEqual([]);
    expect(mapSearchRows(undefined)).toEqual([]);
  });
});
