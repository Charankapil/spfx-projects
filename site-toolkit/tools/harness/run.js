/**
 * Opens every Site Toolkit component in a real browser against the mock
 * SharePoint, clicks through its main paths, and fails on:
 *   - any console error or uncaught exception,
 *   - two requests in flight at once, or a gap shorter than the queue allows,
 *   - an unexpected number of requests,
 *   - list content running as markup,
 *   - any expectation in a scenario not being met.
 *
 * Usage: node tools/harness/run.js [scenario-name-filter] [--shots <dir>]
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ORIGIN = 'https://contoso.sharepoint.com';
const PAGE = `${ORIGIN}/sites/team/SitePages/harness.aspx`;
const bundle = fs.readFileSync(path.join(__dirname, '..', '..', 'temp', 'harness', 'harness.bundle.js'));
const args = process.argv.slice(2);
const shotsIndex = args.indexOf('--shots');
const shotsDir = shotsIndex >= 0 ? args[shotsIndex + 1] : undefined;
const only = args.filter((a, i) => a.indexOf('--') !== 0 && i !== shotsIndex + 1)[0];
// A 1x1 PNG standing in for a profile photo.
const PHOTO = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');
/** The queue leaves 500 ms between requests; timers may fire a few ms early. */
const MIN_GAP_MS = 480;

const scenarios = [];
function scenario(name, params, steps) {
  scenarios.push({ name, params, steps });
}

function url(params) {
  const q = new URLSearchParams();
  for (const key of Object.keys(params)) {
    q.set(key, typeof params[key] === 'string' ? params[key] : JSON.stringify(params[key]));
  }
  return `${PAGE}?${q.toString()}`;
}

class Helper {
  constructor(page, name) {
    this.page = page;
    this.name = name;
    this.failures = [];
  }
  expect(condition, message) {
    if (!condition) {
      this.failures.push(message);
    }
  }
  /** Waits until no request is in flight and none has started for a while. */
  async idle(quietMs = 1100) {
    let last = -1;
    let stableSince = Date.now();
    const deadline = Date.now() + 90000;
    while (Date.now() < deadline) {
      const state = await this.page.evaluate(() => ({ n: window.__sp.log.length, pending: window.__sp.log.filter((e) => !e.end).length }));
      if (state.n !== last || state.pending > 0) {
        last = state.n;
        stableSince = Date.now();
      } else if (Date.now() - stableSince >= quietMs) {
        return;
      }
      await this.page.waitForTimeout(100);
    }
    this.failures.push('the page never became idle');
  }
  async text() {
    return this.page.evaluate(() => document.body.innerText);
  }
  async has(text, message) {
    const body = await this.text();
    this.expect(body.indexOf(text) >= 0, message || `expected to see "${text}"`);
  }
  async hasNot(text, message) {
    const body = await this.text();
    this.expect(body.indexOf(text) < 0, message || `did not expect to see "${text}"`);
  }
  async click(name, role = 'button') {
    await this.page.getByRole(role, { name, exact: false }).first().click();
  }
  async choose(comboLabel, optionText) {
    await this.page.getByRole('combobox', { name: comboLabel }).first().click();
    await this.page.getByRole('option', { name: optionText, exact: true }).first().click();
  }
  /** Types a date the way a person would (the test browser's date fields are month/day/year). */
  async typeDate(label, iso) {
    const [year, month, day] = iso.split('-');
    const input = this.page.getByLabel(label);
    await input.focus();
    await this.page.keyboard.type(`${month}${day}${year}`);
    this.expect((await input.inputValue()) === iso, `typing a date into "${label}" gave ${await input.inputValue()}`);
  }
  async requests() {
    return this.page.evaluate(() => window.__sp.log.map((e) => ({ method: e.method, url: e.url, status: e.status, start: e.start, end: e.end, overlap: e.overlap })));
  }
  async count() {
    return (await this.requests()).length;
  }
  async shot(label) {
    if (shotsDir) {
      await this.page.screenshot({ path: path.join(shotsDir, `${this.name}${label ? `-${label}` : ''}.png`), fullPage: true });
    }
  }
}

// ---------------------------------------------------------------- scenarios

scenario('storage-owner', { wp: 'StorageHogFinder', role: 'owner' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 2, `opening should take 2 requests, took ${await h.count()}`);
  await h.has('Documents');
  await h.has("Finance & Legal's");
  await h.hasNot('_catalogs', 'system folders must not be listed');
  await h.has('Older versions and metadata');
  await h.shot('libraries');
  await h.click('Documents');
  await h.idle();
  h.expect((await h.count()) === 3, 'drilling into a folder should take exactly 1 request');
  await h.has('Folder A');
  await h.hasNot('Forms', 'the Forms folder must be hidden');
  await h.click('Show the files in this folder');
  await h.idle();
  await h.has('Budget 2024.xlsx');
  await h.page.getByRole('button', { name: 'Versions' }).first().click();
  await h.idle();
  await h.has('14 older versions');
  await h.shot('versions');
  await h.click('Trim 4 older versions');
  await h.has('will be moved to the site recycle bin');
  await h.click('Yes, trim them');
  await h.idle();
  await h.has('4 older versions of "Budget 2024.xlsx" moved to the recycle bin');
  const left = await h.page.evaluate(() => window.__sp.state.versions['/sites/team/Shared Documents/Budget 2024.xlsx'].map((v) => v.VersionLabel));
  h.expect(left.length === 10 && left[0] === '5.0', `the 10 newest versions must remain, got ${left.join(',')}`);
  await h.page.getByRole('button', { name: 'Close' }).first().click();
  await h.click('Largest files', 'tab');
  await h.idle();
  await h.has('Copy of report.docx');
  await h.shot('largest');
});

scenario('storage-member', { wp: 'StorageHogFinder', role: 'member' }, async (h) => {
  await h.idle();
  await h.has('only available to people who can manage this site');
});

scenario('readack-user', { wp: 'ReadAcknowledge' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 2, `loading should take 2 requests, took ${await h.count()}`);
  await h.has('2 policies need your acknowledgement');
  await h.has('Updated - please read again');
  await h.hasNot('Who has acknowledged', 'the report is for owners only');
  const ack = h.page.getByRole('button', { name: 'Acknowledge', exact: true });
  h.expect(await ack.first().isDisabled(), 'Acknowledge must be disabled until the box is ticked');
  await h.page.getByRole('checkbox').first().check({ force: true });
  await ack.first().click();
  await h.idle();
  await h.has('1 policy needs your acknowledgement');
  await h.has('Your acknowledgement of "IT Security.docx" is recorded');
  const saved = await h.page.evaluate(() => window.__sp.state.lists['/sites/team/lists/toolkitpolicyacks'].items.slice(-1)[0]);
  h.expect(saved.PolicyId === 3 && saved.PolicyVersion === '1', 'the acknowledgement must record the policy and its major version');
  await h.shot();
});

scenario('readack-owner-report', { wp: 'ReadAcknowledge', role: 'owner', user: '8' }, async (h) => {
  await h.idle();
  await h.click('Who has acknowledged');
  await h.idle();
  await h.has('contains groups as well as people');
  await h.page.getByRole('button', { name: 'Show names' }).first().click();
  await h.has('Alex Reader, Priya Shah');
  await h.has('Tom Okafor');
  await h.shot();
});

scenario('readack-setup-member', { wp: 'ReadAcknowledge', state: 'empty' }, async (h) => {
  await h.idle();
  await h.has('ask a site owner');
  await h.hasNot('Set up');
});

scenario('readack-setup-owner', { wp: 'ReadAcknowledge', state: 'empty', role: 'owner' }, async (h) => {
  await h.idle();
  await h.click('Set up');
  await h.idle();
  await h.has('There are no documents in the "Policies" library yet');
  const lists = await h.page.evaluate(() => Object.keys(window.__sp.state.lists));
  h.expect(lists.indexOf('/sites/team/policies') >= 0 && lists.indexOf('/sites/team/lists/toolkitpolicyacks') >= 0, 'setup must create the library and the list');
  const title = await h.page.evaluate(() => window.__sp.state.lists['/sites/team/lists/toolkitpolicyacks'].Title);
  h.expect(title === 'Policy Acknowledgements', `the list must get its display title, got ${title}`);
});

scenario('stale', { wp: 'StaleContentRadar' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `loading should take 1 request, took ${await h.count()}`);
  await h.has('By file type');
  await h.has('Load 100 more');
  h.expect((await h.page.locator('tbody tr').count()) === 100, 'the first page shows 100 files');
  await h.click('Load 100 more');
  await h.idle();
  h.expect((await h.page.locator('tbody tr').count()) === 200, 'the second page adds 100 files');
  h.expect((await h.count()) === 2, 'loading more should take exactly 1 request');
  await h.shot();
  await h.choose('Stale after', '5 years');
  await h.idle();
  h.expect((await h.page.locator('tbody tr').count()) <= 100, 'changing the period starts a fresh list');
});

scenario('stale-search-down', { wp: 'StaleContentRadar', options: { searchFails: true } }, async (h) => {
  await h.idle();
  await h.has('Stale files could not be found: Search is unavailable.');
});

scenario('stale-throttled', { wp: 'StaleContentRadar', options: { throttleAt: [1] } }, async (h) => {
  await h.page.waitForTimeout(700);
  await h.has('SharePoint asked this page to slow down', 'a throttle must be explained, not hidden');
  await h.shot('paused');
  await h.idle(3500);
  await h.has('By file type', 'the request must be repeated and succeed after the pause');
  await h.hasNot('could not be found');
  const log = await h.requests();
  h.expect(log.length === 2 && log[0].status === 429 && log[1].status === 200, 'exactly one retry after one 429');
  h.expect(log[1].start - log[0].end >= 1950, `the retry must wait for Retry-After (waited ${Math.round(log[1].start - log[0].end)} ms)`);
});

scenario('recycle', { wp: 'RecycleBinExplorer', role: 'owner' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `loading should take 1 request, took ${await h.count()}`);
  h.expect((await h.page.locator('tbody tr').count()) === 10, 'first-stage items only for the site view');
  await h.has('Deleted 3.docx', 'an item with no leaf name falls back to its title');
  await h.page.getByPlaceholder('Filter by name').fill('deleted 5');
  h.expect((await h.page.locator('tbody tr').count()) === 1, 'the text filter narrows the list without a request');
  await h.page.getByPlaceholder('Filter by name').fill('');
  await h.page.getByRole('button', { name: 'Restore' }).nth(2).click();
  await h.idle();
  await h.has('was restored to /sites/team/Shared Documents');
  h.expect((await h.page.locator('tbody tr').count()) === 9, 'a restored item leaves the list');
  await h.page.getByRole('button', { name: 'Restore' }).nth(1).click();
  await h.idle();
  await h.has('could not be restored: A file with this name already exists');
  h.expect(await h.page.getByRole('button', { name: 'Restore' }).nth(1).isEnabled(), 'a failed restore can be tried again');
  await h.shot();
  await h.page.getByRole('switch').click();
  await h.idle();
  await h.has('second stage');
});

scenario('recycle-no-server-sort', { wp: 'RecycleBinExplorer', options: { noRecycleOrder: true } }, async (h) => {
  await h.idle();
  h.expect((await h.page.locator('tbody tr').count()) === 10, 'falls back to sorting in the browser');
  const first = await h.page.locator('tbody tr').first().innerText();
  h.expect(first.indexOf('Deleted 0.docx') >= 0, 'newest deletion first');
  await h.hasNot('Whole site collection', 'the site-wide switch is for owners');
});

scenario('bulk', { wp: 'BulkMetadataEditor', role: 'member' }, async (h) => {
  await h.idle();
  await h.choose('Library', 'Documents');
  await h.idle();
  h.expect((await h.count()) === 3, `libraries + columns + first page should be 3 requests, took ${await h.count()}`);
  await h.has('Due date');
  await h.hasNot('Owner', 'person columns are not offered');
  h.expect((await h.page.locator('tbody tr').count()) === 7, 'folders are not listed');
  const title0 = h.page.getByLabel('Title of Budget 2024.xlsx');
  await title0.fill('Annual budget');
  await h.page.getByLabel('Amount of Budget 2024.xlsx').fill('abc');
  await h.has('1 highlighted value does not fit');
  await title0.focus();
  await h.click('Fill down');
  h.expect((await h.page.getByLabel("Title of Bob's plan.docx").inputValue()) === 'Annual budget', 'fill down copies the value to the rows below');
  await h.click('Undo');
  h.expect((await h.page.getByLabel("Title of Bob's plan.docx").inputValue()) === 'Title 1', 'undo reverts the fill down in one step');
  await h.page.getByLabel('Status of Roadmap.pptx').selectOption('Retired');
  await h.page.getByLabel('Reviewed of Roadmap.pptx').selectOption('No');
  await h.typeDate('Due date of Roadmap.pptx', '2030-01-15');
  await h.shot('editing');
  const before = await h.count();
  await h.page.getByRole('button', { name: /^Save/ }).click();
  await h.idle();
  h.expect((await h.count()) === before + 1, 'all changes travel in a single batch request');
  await h.has('2 files updated');
  const items = await h.page.evaluate(() => window.__sp.state.lists['/sites/team/shared documents'].items);
  h.expect(items[0].Title === 'Annual budget' && items[0].Amount === 0, 'the valid change is saved and the invalid one is not');
  h.expect(items[2].Status === 'Retired' && items[2].Reviewed === false && /^2030-01-1[45]T/.test(items[2].DueDate), `choice, yes/no and date are saved in SharePoint's formats (${items[2].DueDate})`);
  await h.has('1 highlighted value does not fit');
});

scenario('bulk-partial-failure', { wp: 'BulkMetadataEditor', options: { batchFailsItem: 2, noChoicesSelect: true } }, async (h) => {
  await h.idle();
  await h.choose('Library', 'Documents');
  await h.idle();
  await h.has('Status', 'columns still load when selecting Choices is refused');
  await h.page.getByLabel('Title of Budget 2024.xlsx').fill('One');
  await h.page.getByLabel("Title of Bob's plan.docx").fill('Two');
  await h.page.getByRole('button', { name: /^Save/ }).click();
  await h.idle();
  await h.has('1 file updated');
  await h.has('1 file was not saved and remain highlighted');
  await h.has('The file is locked for editing');
  h.expect((await h.page.getByLabel("Title of Bob's plan.docx").inputValue()) === 'Two', 'the failed edit stays in the grid');
  await h.has('Save 1 change');
});

scenario('bulk-paste', { wp: 'BulkMetadataEditor' }, async (h) => {
  await h.idle();
  await h.choose('Library', 'Documents');
  await h.idle();
  await h.page.getByLabel('Title of Budget 2024.xlsx').focus();
  await h.page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'A1\tDraft\r\nA2\tApproved\r\nA3\tNot a choice\r\n');
    document.activeElement.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  h.expect((await h.page.getByLabel("Title of Bob's plan.docx").inputValue()) === 'A2', 'a pasted block fills the rows below');
  h.expect((await h.page.getByLabel('Status of Budget 2024.xlsx').inputValue()) === 'Draft', 'and the columns to the right');
  await h.has('1 highlighted value does not fit');
  await h.click('Undo');
  h.expect((await h.page.getByLabel("Title of Bob's plan.docx").inputValue()) === 'Title 1', 'a paste is undone in one step');
});

scenario('duplicates', { wp: 'DuplicateFileDetector' }, async (h) => {
  await h.idle();
  await h.click('Scan');
  await h.idle(1500);
  await h.has('Scan finished');
  await h.has('Copy of report.docx');
  const search = (await h.requests()).filter((r) => r.url.indexOf('/_api/search/') >= 0);
  h.expect(search.length === 3, `1,234 files at 500 per request is 3 requests, took ${search.length}`);
  await h.has('1,234');
  await h.click('Show copies');
  await h.has('/sites/team/Shared Documents/Folder 0');
  await h.shot();
  const before = await h.count();
  await h.page.reload();
  await h.idle();
  await h.has('Scan finished', 'a finished scan is shown again from the cache');
  h.expect((await h.requests()).filter((r) => r.url.indexOf('/_api/search/') >= 0).length === 0, `a cached result needs no search request (${before})`);
});

scenario('duplicates-pause-resume', { wp: 'DuplicateFileDetector', options: { latency: 400 } }, async (h) => {
  await h.idle();
  await h.click('Scan');
  await h.page.waitForTimeout(700);
  await h.click('Pause');
  await h.idle(1500);
  await h.has('Partial result');
  const paused = (await h.requests()).filter((r) => r.url.indexOf('/_api/search/') >= 0).length;
  await h.click('Resume');
  await h.idle(1500);
  await h.has('Scan finished');
  const total = (await h.requests()).filter((r) => r.url.indexOf('/_api/search/') >= 0).length;
  h.expect(total === 3, `resume must carry on, not start over (${paused} before pause, ${total} in total)`);
  await h.has('1,234');
});

scenario('myfiles', { wp: 'MyFiles' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 2, `loading should take 2 requests, took ${await h.count()}`);
  await h.has('Checked out to me (2)');
  await h.has('Recently edited by me (10)');
  await h.shot();
  await h.page.reload();
  await h.idle();
  h.expect((await h.count()) === 0, 'a reload within 5 minutes uses the cache: 0 requests');
  await h.has('Checked out to me (2)');
});

scenario('myfiles-empty', { wp: 'MyFiles', state: 'empty' }, async (h) => {
  await h.idle();
  await h.has('You have no files checked out');
  await h.has('You have not edited any files');
});

scenario('myfiles-search-down', { wp: 'MyFiles', options: { searchFails: true } }, async (h) => {
  await h.idle();
  await h.has('This list could not be loaded');
});

scenario('poll', { wp: 'PulsePoll' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `an unvoted member needs 1 request, took ${await h.count()}`);
  await h.has('How was your week?');
  await h.hasNot('%', 'results are hidden until you vote');
  await h.click('Good');
  await h.idle();
  await h.has('your answer');
  await h.has('4 votes', 'votes for a removed answer and for other polls are not counted');
  await h.has('50% (2)');
  await h.shot();
  await h.page.reload();
  await h.idle();
  await h.has('your answer');
  h.expect((await h.count()) <= 1, `after voting, a reload needs at most the results request, took ${await h.count()}`);
});

scenario('poll-owner', { wp: 'PulsePoll', role: 'owner' }, async (h) => {
  await h.idle();
  await h.has('Results so far', 'an owner sees results before voting');
  await h.has('Good', 'and can still vote');
});

scenario('poll-already-voted', { wp: 'PulsePoll', user: '8' }, async (h) => {
  await h.idle();
  await h.has('your answer', 'a vote cast earlier (or on another device) is found');
  h.expect((await h.page.locator('[data-wp] button').filter({ hasText: /^Great$/ }).count()) === 0, 'someone who voted cannot vote again');
});

scenario('poll-unconfigured', { wp: 'PulsePoll', props: { question: '', options: 'Only one' } }, async (h) => {
  await h.idle();
  await h.has('enter a question with at least two answers');
  h.expect((await h.count()) === 0, 'an unconfigured poll asks SharePoint nothing');
});

scenario('poll-setup-member', { wp: 'PulsePoll', state: 'empty' }, async (h) => {
  await h.idle();
  await h.has('ask a site owner');
});

scenario('templates', { wp: 'TemplateGallery' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 2, `loading should take 2 requests, took ${await h.count()}`);
  await h.has('Project Brief');
  await h.page.getByRole('button', { name: 'Use this template' }).nth(1).click();
  const name = h.page.getByLabel('Name');
  h.expect((await name.inputValue()) === 'Project Brief', 'the name starts as the template name');
  const options = await h.page.getByRole('combobox', { name: 'Save in' }).innerText();
  h.expect(options.indexOf('Templates') < 0, 'the template library is not offered as a destination');
  await name.fill('bad/name');
  await h.click('Create');
  await h.has('A name cannot contain');
  await name.fill('already exists');
  await h.click('Create');
  await h.idle();
  await h.has('The document could not be created');
  await name.fill("Q3 brief & Bob's notes");
  await h.shot('form');
  await h.click('Create');
  await h.idle();
  await h.has('"Q3 brief & Bob\'s notes.docx" is ready');
  const copy = (await h.requests()).filter((r) => r.url.indexOf('copyTo') >= 0).pop();
  h.expect(copy.url.indexOf("strNewUrl='/sites/team/Shared Documents/Q3 brief & Bob''s notes.docx'") >= 0, `the destination must be quoted safely: ${copy.url}`);
  const href = await h.page.getByRole('link', { name: 'Open it' }).getAttribute('href');
  h.expect(href === "/sites/team/Shared%20Documents/Q3%20brief%20&%20Bob's%20notes.docx?web=1", `the open link must be a valid address: ${href}`);
});

scenario('templates-setup', { wp: 'TemplateGallery', state: 'empty', role: 'owner' }, async (h) => {
  await h.idle();
  await h.click('Set up');
  await h.idle();
  await h.has('There are no templates yet');
});

scenario('faq', { wp: 'FaqAccordion' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `loading should take 1 request, took ${await h.count()}`);
  const order = await h.page.locator('button[aria-controls]').allInnerTexts();
  h.expect(order[0].indexOf('no answer') >= 0 && order[1].indexOf('holiday') >= 0, `questions follow the sort order: ${order.join(' | ')}`);
  await h.click('Where is the holiday form?');
  await h.has('<b>Not bold</b> & safe', 'markup in an answer is shown as text');
  await h.click('Question with no answer or category');
  await h.has('No answer has been written yet');
  await h.page.getByPlaceholder('Search the questions').fill('PASSWORD');
  h.expect((await h.page.locator('button[aria-controls]').count()) === 1, 'search is case-insensitive and needs no request');
  await h.page.getByPlaceholder('Search the questions').fill('');
  await h.click('HR');
  h.expect((await h.page.locator('button[aria-controls]').count()) === 1, 'category filter');
  h.expect((await h.count()) === 1, 'searching and filtering ask SharePoint nothing');
  await h.shot();
  await h.hasNot('Edit questions');
});

scenario('faq-setup', { wp: 'FaqAccordion', state: 'empty', role: 'owner' }, async (h) => {
  await h.idle();
  await h.click('Set up');
  await h.idle(1500);
  await h.has('How do I add a question here?');
  await h.has('Edit questions');
});

scenario('access-owner', { wp: 'WhoHasAccess', role: 'owner' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `loading should take 1 request, took ${await h.count()}`);
  await h.has('External guest');
  await h.has('Everyone in the organisation');
  await h.hasNot('Only limited', 'principals with only Limited Access are not shown');
  const members = await h.page.locator('tr', { hasText: 'Team Members' }).innerText();
  h.expect(members.indexOf('Limited Access') < 0 && members.indexOf('Edit') >= 0, 'Limited Access is not listed as a level');
  await h.click('Team Members');
  await h.idle();
  await h.has('Guest Partner');
  await h.shot('site');
  await h.click('Shared outside', 'tab');
  await h.idle();
  await h.has('Anyone with the link');
  await h.page.getByRole('button', { name: 'Who can open it' }).first().click();
  await h.idle();
  await h.has('Full Control');
  await h.shot('shared');
  await h.click('Own permissions', 'tab');
  await h.idle();
  await h.has("Finance & Legal's");
  await h.has('Tasks');
  const link = await h.page.getByRole('link', { name: 'Review permissions' }).first().getAttribute('href');
  h.expect(/user\.aspx\?obj=%7Blib-fin%7D,doclib&List=%7Blib-fin%7D$/.test(link), `permission link: ${link}`);
});

scenario('access-member', { wp: 'WhoHasAccess', role: 'member' }, async (h) => {
  await h.idle();
  await h.has('can only be seen by people who are allowed to manage permissions');
});

scenario('links', { wp: 'BrokenLinkChecker' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 0, 'nothing is requested until the check is started');
  await h.click('Check links');
  await h.idle(1500);
  await h.has('Checked ');
  const text = await h.text();
  h.expect(/2\s*\n\s*Broken links/.test(text), `two broken links expected:\n${text.substring(0, 400)}`);
  await h.has('/sites/team/missing-page.aspx');
  await h.has('/sites/hr/missing-policy.pdf', 'links kept inside web part data are found too');
  await h.has('Home, Other.aspx', 'a link is checked once and lists every page it is on');
  await h.has('1 link to other websites was not checked');
  const heads = (await h.requests()).filter((r) => r.method === 'HEAD');
  h.expect(heads.length === 5, `5 distinct internal links should be checked once each, checked ${heads.length}: ${heads.map((x) => x.url).join(' ')}`);
  await h.choose('Show', 'All checked links');
  await h.has('No access for you (403)');
  await h.has('Working (200)');
  await h.shot();
});

scenario('alert-settings', { wp: 'AlertBannerSettings', role: 'owner' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 2, `loading should take 2 requests, took ${await h.count()}`);
  await h.has('Planned maintenance');
  await h.has('scheduled');
  await h.has('expired');
  await h.page.getByRole('switch').click();
  await h.idle();
  await h.has('The banner is on for this site collection');
  const action = await h.page.evaluate(() => window.__sp.state.customActions[0]);
  h.expect(action && action.ClientSideComponentId === 'b2be5f7c-564a-4ad3-ba1a-16ab1b12a70e' && action.Location === 'ClientSideExtension.ApplicationCustomizer', 'the banner is registered as an application customizer');
  await h.click('New alert');
  await h.click('Save alert');
  await h.has('Enter a headline');
  await h.page.getByLabel('Headline').fill('Fire drill at 3pm');
  await h.page.getByLabel('Link (optional)').fill('javascript:alert(1)');
  await h.click('Save alert');
  await h.has('The link must start with https://');
  await h.page.getByLabel('Link (optional)').fill('');
  await h.typeDate('Show until (optional)', '2020-01-01');
  await h.click('Save alert');
  await h.has('The end date is before the start date');
  await h.page.getByLabel('Show until (optional)').focus();
  for (let i = 0; i < 3; i++) {
    await h.page.keyboard.press('Backspace');
    await h.page.keyboard.press('Tab');
  }
  h.expect((await h.page.getByLabel('Show until (optional)').inputValue()) === '', 'the end date can be cleared again');
  await h.shot('form');
  await h.click('Save alert');
  await h.idle();
  await h.has('Fire drill at 3pm');
  await h.page.getByRole('button', { name: 'Remove' }).first().click();
  await h.idle();
  h.expect((await h.page.locator('tbody').innerText()).indexOf('Fire drill at 3pm') < 0, 'a removed alert leaves the table');
  await h.has('"Fire drill at 3pm" was moved to the recycle bin');
  await h.page.getByRole('switch').click();
  await h.idle();
  await h.has('The banner is off');
  h.expect((await h.page.evaluate(() => window.__sp.state.customActions.length)) === 0, 'switching off removes the registration');
  await h.shot();
});

scenario('alert-settings-member', { wp: 'AlertBannerSettings' }, async (h) => {
  await h.idle();
  await h.has('Only site owners can manage the alert banner');
  h.expect((await h.count()) === 0, 'a member causes no requests');
});

scenario('alert-settings-not-admin', { wp: 'AlertBannerSettings', role: 'owner', options: { notSiteAdmin: true } }, async (h) => {
  await h.idle();
  await h.has('Only a site collection administrator can see or change');
  h.expect(await h.page.getByRole('switch').isDisabled(), 'the switch is disabled when its state is unknown');
});

scenario('alert-settings-setup', { wp: 'AlertBannerSettings', role: 'owner', state: 'empty' }, async (h) => {
  await h.idle();
  await h.click('Set up');
  await h.idle(1500);
  await h.has('There are no alerts yet');
});

scenario('banner', { wp: 'AlertBanner' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, `the banner must ask once at page load, asked ${await h.count()} times`);
  const banners = await h.page.locator('#top-placeholder > div').allInnerTexts();
  h.expect(banners.length === 2, `two alerts are active, showing ${banners.length}`);
  h.expect(banners[0].indexOf('Office closed') >= 0, 'critical alerts come first');
  h.expect(banners[1].indexOf('<script>alert(1)</script>') >= 0, 'markup in an alert is shown as text');
  const links = await h.page.locator('#top-placeholder a').evaluateAll((a) => a.map((x) => x.getAttribute('href')));
  h.expect(links.length === 1 && links[0] === 'https://example.com/status', `only the safe link is rendered: ${links.join(',')}`);
  h.expect((await h.page.locator('#top-placeholder [role=alert]').count()) === 1, 'critical alerts are announced');
  await h.shot();
  await h.page.evaluate(() => window.__navigate());
  await h.idle();
  h.expect((await h.count()) === 1, 'navigating inside the site draws from the cache');
  h.expect((await h.page.locator('#top-placeholder > div').count()) === 2, 'redrawing does not duplicate banners');
  await h.page.getByRole('button', { name: 'Dismiss: Office closed' }).click();
  h.expect((await h.page.locator('#top-placeholder > div').count()) === 1, 'a dismissed alert goes away');
  await h.page.reload();
  await h.idle();
  h.expect((await h.count()) === 0, 'the next page load within 15 minutes asks SharePoint nothing');
  h.expect((await h.page.locator('#top-placeholder > div').count()) === 1, 'a dismissed alert stays dismissed');
});

scenario('banner-no-list', { wp: 'AlertBanner', state: 'empty' }, async (h) => {
  await h.idle();
  h.expect((await h.count()) === 1, 'one request finds out there is no list');
  h.expect((await h.page.locator('#top-placeholder > div').count()) === 0, 'nothing is drawn');
  await h.page.reload();
  await h.idle();
  h.expect((await h.count()) === 0, 'the missing list is remembered instead of asked for on every page');
});

scenario('org', { wp: 'OrgChart' }, async (h) => {
  await h.idle(1500);
  const log = await h.requests();
  const rest = log.filter((r) => r.url.indexOf('userphoto') < 0);
  h.expect(rest.length === 2, `opening needs 1 profile request and 1 batch, took ${rest.length}: ${rest.map((r) => r.url.slice(-60)).join(' ')}`);
  h.expect(rest[1].url.endsWith('/_api/$batch'), 'managers and reports come in one batch');
  await h.has('Alex Reader');
  await h.has('(you)');
  await h.has('Dana Whitfield', 'two manager levels are shown');
  await h.has('Priya Shah');
  await h.has('Direct reports (2)', 'the count matches the people shown');
  await h.has('1 more listed as reporting here has no profile');
  await h.has('2 reports', 'manager cards say how many reports they have');
  await h.has('Jo Nguyen');
  await h.has('Software Engineer <b>Intern</b>', 'markup in a job title is shown as text');
  await h.hasNot('gone@', 'a report without a profile is left out');
  h.expect((await h.page.locator('a[href^="tel:"]').getAttribute('href')) === 'tel:+61255501234', 'the phone link is dialable');
  const photos = log.filter((r) => r.url.indexOf('userphoto') >= 0);
  h.expect(photos.length === 5, `each of the 5 people shown gets one photo request, got ${photos.length}`);
  h.expect((await h.page.locator('img').count()) === 2, 'people with a photo show it, the others show initials');
  await h.shot();

  await h.click('Priya Shah');
  await h.idle(1500);
  await h.has('Head of Engineering');
  await h.has('Direct reports (2)');
  await h.has('Tom Okafor');
  const after = (await h.requests()).filter((r) => r.url.indexOf('userphoto') < 0);
  h.expect(after.length === 3, `Priya is already known, so moving to her needs only 1 batch for Tom; total now ${after.length}`);
  const photoCalls = (await h.requests()).filter((r) => r.url.indexOf('userphoto') >= 0).length;
  h.expect(photoCalls === 6, `photos already loaded are not fetched again (${photoCalls} photo requests)`);

  await h.click('Back');
  await h.idle(1500);
  await h.has('(you)');
  const back = (await h.requests()).filter((r) => r.url.indexOf('userphoto') < 0).length;

  h.expect(back === 3, `going back to someone already seen costs nothing (${back} requests in total)`);

  await h.page.getByRole('switch').click();
  await h.idle(1500);
  await h.has('Peers (1)');
  await h.shot('peers');

  await h.page.getByLabel('Find a person').pressSequentially('sam', { delay: 60 });
  await h.idle();
  await h.page.getByRole('option', { name: /Sam O'Brien/ }).click();
  await h.idle(1500);
  await h.has('Head of Sales');
  await h.has('No direct reports');
  await h.has('Peers (1)', 'peers stay on after moving');
  await h.click('Me');
  await h.idle(1500);
  await h.has('(you)');
});

scenario('org-more-levels', { wp: 'OrgChart', user: '7', props: { managerLevels: '1' } }, async (h) => {
  await h.idle(1500);
  await h.hasNot('Dana Whitfield');
  await h.click('Show 1 more level above');
  await h.idle(1500);
  await h.has('Dana Whitfield');
  await h.hasNot('more level', 'nothing is hidden any more');
});

scenario('org-no-photos', { wp: 'OrgChart', props: { showPhotos: false } }, async (h) => {
  await h.idle(1500);
  h.expect((await h.requests()).filter((r) => r.url.indexOf('userphoto') >= 0).length === 0, 'with photos off, none are fetched');
  await h.has('AR', 'initials are shown instead');
});

scenario('org-no-profile', { wp: 'OrgChart', options: { noProfile: true } }, async (h) => {
  await h.idle();
  await h.has('No profile was found for this person');
  h.expect((await h.count()) === 1, 'one request, and nothing more');
});

scenario('org-batch-throttled', { wp: 'OrgChart', props: { showPhotos: false }, options: { batchThrottleOnce: true } }, async (h) => {
  await h.idle(3500);
  await h.has('Jo Nguyen', 'a throttled batch is repeated after the wait');
  const log = await h.requests();
  h.expect(log.length === 3, `profile, throttled batch, repeated batch: ${log.length} requests`);
  h.expect(log[2].start - log[1].end >= 1950, 'the repeat waited for Retry-After');
});

scenario('org-profiles-down', { wp: 'OrgChart', options: { forbidden: 'PeopleManager' } }, async (h) => {
  await h.idle();
  await h.has('The org chart could not be loaded: Access denied.');
});

scenario('shared-queue', { wp: 'FaqAccordion,OrgChart,PulsePoll,MyFiles,AlertBanner' }, async (h) => {
  await h.idle(1500);
  const rest = (await h.requests()).filter((r) => r.url.indexOf('userphoto') < 0);
  h.expect(rest.length === 7, `five components should make 7 requests between them, made ${rest.length}`);
  await h.has('How do I reset my password?');
  await h.has('Direct reports (2)');
  await h.has('How was your week?');
});

// ------------------------------------------------------------------- runner

(async () => {
  if (shotsDir) {
    fs.mkdirSync(shotsDir, { recursive: true });
  }
  const browser = await chromium.launch();
  const selected = scenarios.filter((s) => !only || s.name.indexOf(only) >= 0);
  let failed = 0;
  let totalRequests = 0;

  for (const s of selected) {
    const context = await browser.newContext({ viewport: { width: 1100, height: 800 }, timezoneId: 'Australia/Sydney', locale: 'en-GB' });
    await context.route('**/*', (route) => {
      const request = route.request();
      const target = request.url();
      if (target.indexOf(PAGE) === 0) {
        return route.fulfill({ contentType: 'text/html', body: '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Harness</title></head><body style="margin:16px"><div id="root"></div><script src="/harness.bundle.js"></script></body></html>' });
      }
      if (target === `${ORIGIN}/harness.bundle.js`) {
        return route.fulfill({ contentType: 'application/javascript', body: bundle });
      }
      if (target.indexOf('/_layouts/15/userphoto.aspx') >= 0) {
        const withPhoto = /accountname=(dana|priya)/.test(target);
        return withPhoto
          ? route.fulfill({ status: 200, contentType: 'image/png', body: PHOTO })
          : route.fulfill({ status: 404, contentType: 'text/html', body: 'not found' });
      }
      if (request.method() === 'HEAD' && target.indexOf(ORIGIN) === 0) {
        return route.fulfill({ status: /missing/.test(target) ? 404 : /secret/.test(target) ? 403 : 200, body: '' });
      }
      return route.abort();
    });
    const page = await context.newPage();
    const problems = [];
    page.on('console', (message) => {
      // Aborted font and icon downloads are the harness's own doing.
      if (message.type() === 'error' && !/Failed to load resource|net::ERR_FAILED/.test(message.text())) {
        problems.push(`console error: ${message.text()}`);
      }
    });
    page.on('pageerror', (err) => problems.push(`uncaught exception: ${err.message}`));

    const h = new Helper(page, s.name);
    try {
      await page.goto(url(s.params));
      await s.steps(h);
      const log = await h.requests();
      totalRequests += log.length;
      if (await page.evaluate(() => window.__xss)) {
        problems.push('list content ran as script');
      }
      for (let i = 0; i < log.length; i++) {
        if (log[i].overlap) {
          problems.push(`request ${i} started while another was in flight: ${log[i].url}`);
        }
        if (i > 0 && log[i].start - log[i - 1].end < MIN_GAP_MS) {
          problems.push(`only ${Math.round(log[i].start - log[i - 1].end)} ms between requests ${i - 1} and ${i}`);
        }
        if (log[i].status >= 500) {
          problems.push(`request answered ${log[i].status}: ${log[i].method} ${log[i].url}`);
        }
      }
    } catch (err) {
      problems.push(`scenario crashed: ${String(err && err.message).split('\n').slice(0, 6).join(' | ')}`);
      await h.shot('crash').catch(() => undefined);
    }
    const all = h.failures.concat(problems).filter((p) => !(s.params.options && s.params.options.searchFails && /answered 500/.test(p)));
    if (all.length > 0) {
      failed++;
      console.log(`FAIL ${s.name}`);
      all.forEach((p) => console.log(`     - ${p}`));
    } else {
      console.log(`ok   ${s.name}`);
    }
    await context.close();
  }
  await browser.close();
  console.log(`\n${selected.length - failed} of ${selected.length} scenarios passed; ${totalRequests} requests checked for overlap and spacing.`);
  process.exit(failed > 0 ? 1 : 0);
})();
