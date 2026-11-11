import assert from 'node:assert/strict';
import { test } from 'node:test';

import { classify, normalizeLink } from '../src/webparts/brokenLinkChecker/links';
import { display, IColumn, parse } from '../src/webparts/bulkMetadataEditor/cells';
import { addToGroups, collectDuplicates, IDuplicateGroup } from '../src/webparts/duplicateFileDetector/duplicates';
import { parseOptions, tally } from '../src/webparts/pulsePoll/poll';
import { stem, validateFileName } from '../src/webparts/templateGallery/fileNames';
import { flagOf } from '../src/webparts/whoHasAccess/principals';

const ORIGIN = 'https://contoso.sharepoint.com';
const PAGE = `${ORIGIN}/sites/team/SitePages/Home.aspx`;

test('links are resolved against the page and sorted into internal and external', () => {
  assert.deepEqual(normalizeLink('/sites/hr/Docs/a.pdf#page=2', PAGE, ORIGIN), { url: `${ORIGIN}/sites/hr/Docs/a.pdf`, internal: true });
  assert.deepEqual(normalizeLink('Other.aspx', PAGE, ORIGIN), { url: `${ORIGIN}/sites/team/SitePages/Other.aspx`, internal: true });
  assert.deepEqual(normalizeLink('https://CONTOSO.sharepoint.com/x', PAGE, ORIGIN), { url: `${ORIGIN}/x`, internal: true });
  assert.equal(normalizeLink('https://example.com/x', PAGE, ORIGIN)?.internal, false);
  assert.equal(normalizeLink('https://contoso.sharepoint.com.evil.example/x', PAGE, ORIGIN)?.internal, false);
});

test('things that are not web addresses are not checked', () => {
  for (const href of ['', '#top', 'mailto:a@b.c', 'tel:123', 'javascript:void(0)', 'ms-word:ofe|u|https://x', 'data:text/plain,x']) {
    assert.equal(normalizeLink(href, PAGE, ORIGIN), undefined, href);
  }
});

test('only a definite "gone" counts as broken', () => {
  assert.equal(classify(404), 'broken');
  assert.equal(classify(410), 'broken');
  assert.equal(classify(200), 'ok');
  assert.equal(classify(302), 'ok');
  assert.equal(classify(403), 'restricted');
  assert.equal(classify(0), 'unknown');
  assert.equal(classify(500), 'unknown');
  assert.equal(classify(405), 'unknown');
});

const col = (kind: IColumn['kind'], choices: string[] = []): IColumn => ({ name: 'F', title: 'F', kind, choices });

test('cell text is turned into the value SharePoint stores', () => {
  assert.deepEqual(parse(' hello ', col('Text')), { ok: true, value: 'hello' });
  assert.deepEqual(parse('', col('Text')), { ok: true, value: '' });
  assert.deepEqual(parse('', col('Number')), { ok: true, value: null });
  assert.deepEqual(parse('12.5', col('Number')), { ok: true, value: 12.5 });
  assert.equal(parse('twelve', col('Number')).ok, false);
  assert.deepEqual(parse('YES', col('Boolean')), { ok: true, value: true });
  assert.deepEqual(parse('no', col('Boolean')), { ok: true, value: false });
  assert.equal(parse('maybe', col('Boolean')).ok, false);
  assert.deepEqual(parse('approved', col('Choice', ['Draft', 'Approved'])), { ok: true, value: 'Approved' });
  assert.equal(parse('Rejected', col('Choice', ['Draft', 'Approved'])).ok, false);
  assert.equal(parse('2024-13-45', col('DateTime')).ok, false, 'an impossible date is rejected, not rolled over');
  assert.equal(parse('2023-02-29', col('DateTime')).ok, false);
  assert.equal(parse('yesterday', col('DateTime')).ok, false);
  assert.equal(parse('x'.repeat(256), col('Text')).ok, false);
});

test('stored values are shown the way they are edited, and round-trip', () => {
  assert.equal(display(null, 'Text'), '');
  assert.equal(display(true, 'Boolean'), 'Yes');
  assert.equal(display(false, 'Boolean'), 'No');
  assert.equal(display(0, 'Number'), '0');
  const stored = parse('2024-03-09', col('DateTime')).value;
  assert.equal(display(stored, 'DateTime'), '2024-03-09');
});

test('duplicates are grouped by name and size', () => {
  const groups: Record<string, IDuplicateGroup> = {};
  addToGroups(groups, 'https://x/a/Report.docx', 100, 'd', 'u');
  addToGroups(groups, 'https://x/b/REPORT.docx', 100, 'd', 'u');
  addToGroups(groups, 'https://x/c/Report.docx', 101, 'd', 'u');
  addToGroups(groups, 'https://x/a/Report.docx', 100, 'd', 'u');
  addToGroups(groups, 'https://x/a/empty.txt', 0, 'd', 'u');
  addToGroups(groups, 'https://x/b/empty.txt', 0, 'd', 'u');
  addToGroups(groups, 'https://x/a/big.zip', 5000, 'd', 'u');
  addToGroups(groups, 'https://x/b/big.zip', 5000, 'd', 'u');
  const found = collectDuplicates(groups);
  assert.equal(found.length, 2);
  assert.equal(found[0].name, 'big.zip', 'the set wasting most space comes first');
  assert.equal(found[1].copies.length, 2, 'the same path listed twice counts once');
});

test('a file named like an object property does not break grouping', () => {
  const groups: Record<string, IDuplicateGroup> = {};
  addToGroups(groups, 'https://x/a/constructor', 10, 'd', 'u');
  addToGroups(groups, 'https://x/b/constructor', 10, 'd', 'u');
  addToGroups(groups, 'https://x/a/__proto__', 10, 'd', 'u');
  addToGroups(groups, 'https://x/b/__proto__', 10, 'd', 'u');
  assert.equal(collectDuplicates(groups).length, 2);
});

test('poll answers are cleaned up and only real answers are counted', () => {
  assert.deepEqual(parseOptions(' Yes \n\nNo\nyes\n  \nMaybe'), ['Yes', 'No', 'Maybe']);
  assert.deepEqual(parseOptions(''), []);
  assert.equal(parseOptions(Array.from({ length: 30 }, (_, i) => `o${i}`).join('\n')).length, 10);
  assert.deepEqual(tally(['Yes', 'No'], ['Yes', 'Yes', 'No', 'Deleted answer', 'toString']), { Yes: 2, No: 1 });
});

test('new document names are checked before SharePoint is asked', () => {
  assert.equal(validateFileName('Q3 report'), undefined);
  assert.ok(validateFileName('   '));
  assert.ok(validateFileName('a/b'));
  assert.ok(validateFileName('what?'));
  assert.ok(validateFileName('.hidden'));
  assert.ok(validateFileName('trailing.'));
  assert.ok(validateFileName('x'.repeat(121)));
  assert.equal(stem('Contract.v2.docx'), 'Contract.v2');
  assert.equal(stem('README'), 'README');
});

test('guests and "everyone" are flagged', () => {
  assert.equal(flagOf('i:0#.f|membership|jane_gmail.com#ext#@contoso.onmicrosoft.com'), 'external');
  assert.equal(flagOf('c:0-.f|rolemanager|spo-grid-all-users/abc'), 'everyone');
  assert.equal(flagOf('c:0(.s|true'), 'everyone');
  assert.equal(flagOf('i:0#.f|membership|bob@contoso.com'), undefined);
  assert.equal(flagOf(''), undefined);
});
