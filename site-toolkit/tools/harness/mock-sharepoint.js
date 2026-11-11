/**
 * An in-memory stand-in for SharePoint, used only by the browser test
 * harness. It answers the REST calls the web parts make, records every
 * request with its timing, and can throttle on demand.
 *
 * It checks the web parts' own logic and request behaviour. It is not proof
 * that real SharePoint answers in the same shapes.
 */
'use strict';

const ORIGIN = 'https://contoso.sharepoint.com';
const WEB = '/sites/team';

function daysAgo(n) {
  return new Date(Date.now() - n * 86400000).toISOString();
}

function createState(options) {
  const seeded = options.state !== 'empty';
  const state = {
    nextId: 1000,
    lists: {},
    libraries: [
      { Id: 'lib-docs', Title: 'Documents', ItemCount: 7, path: `${WEB}/Shared Documents`, BaseTemplate: 101, HasUniqueRoleAssignments: false },
      { Id: 'lib-fin', Title: "Finance & Legal's", ItemCount: 3, path: `${WEB}/Finance`, BaseTemplate: 101, HasUniqueRoleAssignments: true }
    ],
    recycle: [],
    customActions: [],
    versions: {}
  };

  const addList = (path, title, items, extra) => {
    state.lists[path.toLowerCase()] = Object.assign({ Id: `list-${Object.keys(state.lists).length + 1}`, Title: title, path, items: items || [], fields: ['Title'], seeded: true }, extra || {});
    return state.lists[path.toLowerCase()];
  };

  // Document libraries with editable columns.
  addList(`${WEB}/Shared Documents`, 'Documents', [], { Id: 'lib-docs', library: true });
  addList(`${WEB}/Finance`, "Finance & Legal's", [], { Id: 'lib-fin', library: true });
  const docs = state.lists[`${WEB}/shared documents`];
  const names = ['Budget 2024.xlsx', "Bob's plan.docx", 'Roadmap.pptx', 'Contract.pdf', 'Notes.txt', '=SUM(A1).csv', 'Über résumé.docx'];
  names.forEach((name, i) => {
    docs.items.push({
      Id: i + 1, FileLeafRef: name, FileRef: `${WEB}/Shared Documents/${name}`, FSObjType: 0, Title: i % 2 ? `Title ${i}` : null,
      Status: i % 3 === 0 ? 'Draft' : 'Approved', Amount: i === 2 ? null : i * 10.5, Reviewed: i % 2 === 0, DueDate: i === 1 ? null : daysAgo(-i),
      Modified: daysAgo(i * 40), OData__UIVersionString: `${i + 1}.0`
    });
  });
  docs.items.push({ Id: 50, FileLeafRef: 'Archive', FileRef: `${WEB}/Shared Documents/Archive`, FSObjType: 1 });
  docs.fieldDefs = [
    { InternalName: 'Title', EntityPropertyName: 'Title', Title: 'Title', TypeAsString: 'Text' },
    { InternalName: 'Status', EntityPropertyName: 'Status', Title: 'Status', TypeAsString: 'Choice', Choices: ['Draft', 'Approved', 'Retired'] },
    { InternalName: 'Amount', EntityPropertyName: 'Amount', Title: 'Amount', TypeAsString: 'Number' },
    { InternalName: 'Reviewed', EntityPropertyName: 'Reviewed', Title: 'Reviewed', TypeAsString: 'Boolean' },
    { InternalName: 'DueDate', EntityPropertyName: 'DueDate', Title: 'Due date', TypeAsString: 'DateTime' },
    { InternalName: 'Owner', EntityPropertyName: 'OwnerId', Title: 'Owner', TypeAsString: 'User' },
    { InternalName: 'FileLeafRef', EntityPropertyName: 'FileLeafRef', Title: 'Name', TypeAsString: 'File' }
  ];
  state.lists[`${WEB}/finance`].fieldDefs = [{ InternalName: 'Title', EntityPropertyName: 'Title', Title: 'Title', TypeAsString: 'Text' }];

  // Site pages, for the link checker.
  const link = (href, text) => `<a href="${href}">${text}</a>`;
  addList(`${WEB}/SitePages`, 'Site Pages', [
    {
      Id: 1, Title: 'Home', FileLeafRef: 'Home.aspx', FileRef: `${WEB}/SitePages/Home.aspx`, FSObjType: 0,
      CanvasContent1:
        `<div><p>${link('/sites/team/Shared%20Documents/Budget%202024.xlsx', 'Budget')} ${link('/sites/team/missing-page.aspx', 'Old')} ` +
        `${link('https://example.com/outside', 'Outside')} ${link('mailto:a@b.c', 'Mail')} ${link('#top', 'Top')} ${link('Other.aspx', 'Relative')}</p>` +
        `<div data-sp-webpartdata='${JSON.stringify({ serverProcessedContent: { links: { 'items[0].sourceItem.url': '/sites/hr/missing-policy.pdf', 'items[1].sourceItem.url': '/sites/team/secret/restricted.docx' } } })}'></div></div>`
    },
    { Id: 2, Title: null, FileLeafRef: 'Other.aspx', FileRef: `${WEB}/SitePages/Other.aspx`, FSObjType: 0, CanvasContent1: `<p>${link('/sites/team/missing-page.aspx', 'Old again')}</p>` },
    { Id: 3, Title: 'Empty', FileLeafRef: 'Empty.aspx', FileRef: `${WEB}/SitePages/Empty.aspx`, FSObjType: 0, CanvasContent1: null }
  ], { Id: 'lib-pages', BaseTemplate: 119 });

  if (seeded) {
    addList(`${WEB}/Policies`, 'Policies', [
      { Id: 1, FileLeafRef: 'Code of Conduct.pdf', FileRef: `${WEB}/Policies/Code of Conduct.pdf`, FSObjType: 0, Modified: daysAgo(30), OData__UIVersionString: '3.0' },
      { Id: 2, FileLeafRef: 'Travel Policy.docx', FileRef: `${WEB}/Policies/Travel Policy.docx`, FSObjType: 0, Modified: daysAgo(3), OData__UIVersionString: '2.1' },
      { Id: 3, FileLeafRef: 'IT Security.docx', FileRef: `${WEB}/Policies/IT Security.docx`, FSObjType: 0, Modified: daysAgo(300), OData__UIVersionString: '1.0' }
    ], { library: true });
    addList(`${WEB}/Lists/ToolkitPolicyAcks`, 'Policy Acknowledgements', [
      { Id: 1, Title: 'Code of Conduct.pdf', PolicyId: 1, PolicyVersion: '3', Created: daysAgo(10), AuthorId: 7 },
      { Id: 2, Title: 'Travel Policy.docx', PolicyId: 2, PolicyVersion: '1', Created: daysAgo(100), AuthorId: 7 },
      { Id: 3, Title: 'Code of Conduct.pdf', PolicyId: 1, PolicyVersion: '3', Created: daysAgo(5), AuthorId: 8 }
    ]);
    addList(`${WEB}/Templates`, 'Templates', [
      { Id: 1, FileLeafRef: 'Project Brief.docx', FileRef: `${WEB}/Templates/Project Brief.docx`, FSObjType: 0, Modified: daysAgo(20) },
      { Id: 2, FileLeafRef: 'Budget Sheet.xlsx', FileRef: `${WEB}/Templates/Budget Sheet.xlsx`, FSObjType: 0, Modified: daysAgo(200) }
    ], { library: true });
    state.libraries.push({ Id: 'lib-pol', Title: 'Policies', ItemCount: 3, path: `${WEB}/Policies`, BaseTemplate: 101 });
    state.libraries.push({ Id: 'lib-tpl', Title: 'Templates', ItemCount: 2, path: `${WEB}/Templates`, BaseTemplate: 101 });
    addList(`${WEB}/Lists/ToolkitPollVotes`, 'Poll Votes', [
      { Id: 1, Title: 'Great', PollKey: 'weekly', AuthorId: 8 }, { Id: 2, Title: 'Good', PollKey: 'weekly', AuthorId: 9 },
      { Id: 3, Title: 'Great', PollKey: 'weekly', AuthorId: 10 }, { Id: 4, Title: 'Removed answer', PollKey: 'weekly', AuthorId: 11 },
      { Id: 5, Title: 'Great', PollKey: 'other poll', AuthorId: 12 }
    ]);
    addList(`${WEB}/Lists/ToolkitFaq`, 'FAQ', [
      { Id: 1, Title: 'How do I reset my password?', Answer: 'Go to the portal.\nThen follow the steps.', Category: 'IT', SortOrder: 2 },
      { Id: 2, Title: 'Where is the holiday form?', Answer: '<b>Not bold</b> & safe', Category: 'HR', SortOrder: 1 },
      { Id: 3, Title: 'Question with no answer or category', Answer: null, Category: null, SortOrder: null }
    ]);
    addList(`${WEB}/Lists/ToolkitAlerts`, 'Site Alerts', [
      { Id: 1, Title: 'Planned maintenance', Message: 'SharePoint is read-only on Saturday. <script>alert(1)</script>', Severity: 'Warning', StartDate: daysAgo(1), EndDate: daysAgo(-5), LinkUrl: 'https://example.com/status', Modified: daysAgo(1) },
      { Id: 2, Title: 'Office closed', Message: null, Severity: 'Critical', StartDate: daysAgo(2), EndDate: null, LinkUrl: 'javascript:alert(1)', Modified: daysAgo(2) },
      { Id: 3, Title: 'Old news', Message: 'Expired', Severity: 'Info', StartDate: daysAgo(30), EndDate: daysAgo(20), LinkUrl: null, Modified: daysAgo(30) },
      { Id: 4, Title: 'Coming soon', Message: 'Scheduled', Severity: 'Info', StartDate: daysAgo(-3), EndDate: null, LinkUrl: null, Modified: daysAgo(1) }
    ]);
    state.recycle = Array.from({ length: 12 }, (_, i) => ({
      Id: `rb-${i}`, Title: `Deleted ${i}.docx`, LeafName: i === 3 ? '' : `Deleted ${i}.docx`, DirName: 'sites/team/Shared Documents', ItemType: i % 4 === 0 ? 5 : 1,
      ItemState: i > 9 ? 2 : 1, Size: String(i * 100000), DeletedDate: daysAgo(i * 9), DeletedByName: i % 2 ? 'Priya Shah' : 'Tom Okafor'
    }));
    state.versions[`${WEB}/Shared Documents/Budget 2024.xlsx`] = Array.from({ length: 14 }, (_, i) => ({ ID: 512 * (i + 1), VersionLabel: `${i + 1}.0`, Size: String(2000000 + i * 1000), Created: daysAgo(100 - i) }));
  }

  state.users = {
    7: { Id: 7, Title: 'Alex Reader', EMail: 'alex@contoso.com', Email: 'alex@contoso.com', LoginName: 'i:0#.f|membership|alex@contoso.com', PrincipalType: 1 },
    8: { Id: 8, Title: 'Priya Shah', EMail: 'priya@contoso.com', Email: 'priya@contoso.com', LoginName: 'i:0#.f|membership|priya@contoso.com', PrincipalType: 1 },
    9: { Id: 9, Title: 'Tom Okafor', EMail: 'tom@contoso.com', Email: 'tom@contoso.com', LoginName: 'i:0#.f|membership|tom@contoso.com', PrincipalType: 1 },
    10: { Id: 10, Title: 'Guest Partner', EMail: 'g@partner.com', Email: 'g@partner.com', LoginName: 'i:0#.f|membership|g_partner.com#ext#@contoso.onmicrosoft.com', PrincipalType: 1 },
    11: { Id: 11, Title: 'Sales Team', LoginName: 'c:0t.c|tenant|abc', PrincipalType: 4 }
  };
  // A small organisation for the org chart. "gone" left the company: listed as a report, but without a profile.
  const acct = (n) => `i:0#.f|membership|${n}@contoso.com`;
  const person = (n, name, title, dept, managers, reports, peers, phone) => ({
    AccountName: acct(n), DisplayName: name, Email: `${n}@contoso.com`, Title: title,
    ExtendedManagers: managers.map(acct), DirectReports: reports.map(acct), Peers: peers.map(acct),
    UserProfileProperties: [{ Key: 'Department', Value: dept }, { Key: 'Office', Value: 'Sydney' }, { Key: 'WorkPhone', Value: phone || null }]
  });
  state.org = {};
  [
    person('dana', 'Dana Whitfield', 'Chief Executive', 'Leadership', [], ['priya', 'sam'], []),
    person('priya', 'Priya Shah', 'Head of Engineering', 'Engineering', ['dana'], ['alex', 'tom'], ['sam']),
    person('sam', "Sam O'Brien", 'Head of Sales', 'Sales', ['dana'], [], ['priya']),
    person('alex', 'Alex Reader', 'Engineering Manager', 'Engineering', ['dana', 'priya'], ['jo', 'lee', 'gone'], ['tom'], '+61 2 5550 1234'),
    person('tom', 'Tom Okafor', 'Engineering Manager', 'Engineering', ['dana', 'priya'], [], ['alex']),
    person('jo', 'Jo Nguyen', 'Software Engineer', 'Engineering', ['dana', 'priya', 'alex'], [], ['lee']),
    person('lee', 'Lee Fischer', 'Software Engineer <b>Intern</b>', 'Engineering', ['dana', 'priya', 'alex'], [], ['jo'])
  ].forEach((p) => {
    state.org[p.AccountName.toLowerCase()] = p;
  });
  state.users[20] = { Id: 20, Title: 'Dana Whitfield', Email: 'dana@contoso.com', LoginName: acct('dana'), PrincipalType: 1 };
  state.users[21] = { Id: 21, Title: "Sam O'Brien", Email: 'sam@contoso.com', LoginName: acct('sam'), PrincipalType: 1 };
  return state;
}

const SEARCH_FILES = Array.from({ length: 1234 }, (_, i) => ({
  DocId: String(17592186044416 + i * 3),
  Path: `${ORIGIN}${WEB}/${i % 5 === 0 ? 'Finance' : 'Shared Documents'}/${i % 7 === 0 ? 'Copy of report.docx' : `File ${i}.docx`}`,
  Size: String(i % 7 === 0 ? 250000 : 1000 + i * 37),
  LastModifiedTime: daysAgo(i * 3),
  ModifiedBy: i % 2 ? 'Priya Shah' : 'Tom Okafor',
  ViewableByAnonymousUsers: i % 9 === 0 ? 'true' : 'false'
}));
// Give the "Copy of report" files distinct folders so they are real duplicates.
SEARCH_FILES.forEach((f, i) => {
  if (i % 7 === 0) {
    f.Path = `${ORIGIN}${WEB}/Shared Documents/Folder ${i}/Copy of report.docx`;
  }
});

function json(status, body, headers) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  const map = headers || {};
  return {
    raw: text,
    ok: status >= 200 && status < 300,
    status,
    statusText: String(status),
    headers: { get: (name) => (Object.keys(map).filter((k) => k.toLowerCase() === name.toLowerCase()).map((k) => map[k])[0] || null) },
    text: () => Promise.resolve(text),
    json: () => Promise.resolve(JSON.parse(text))
  };
}

const notFound = () => json(404, { error: { code: '-2130575338', message: 'File Not Found.' } });
const denied = () => json(403, { error: { code: '-2147024891', message: 'Access denied.' } });

function quoted(url, marker) {
  const index = url.indexOf(marker);
  if (index < 0) {
    return undefined;
  }
  const start = index + marker.length;
  // The closing quote is the first single quote not doubled.
  let end = start;
  while (end < url.length) {
    if (url.charAt(end) === "'" || url.substr(end, 3) === '%27') {
      break;
    }
    end++;
  }
  return decodeURIComponent(url.substring(start, end)).replace(/''/g, "'");
}

function param(url, name) {
  const match = new RegExp(`[?&]${name.replace('$', '\\$')}=([^&]*)`).exec(url);
  return match ? decodeURIComponent(match[1]) : undefined;
}

function createMock(options) {
  const state = createState(options);
  const me = options.userId || 7;
  const owner = options.role === 'owner';
  const log = [];
  let active = 0;
  let count = 0;
  const latency = options.latency === undefined ? 30 : options.latency;

  function expandUsers(item) {
    const out = Object.assign({}, item);
    if (item.AuthorId) {
      out.Author = { Title: (state.users[item.AuthorId] || { Title: 'Someone' }).Title };
    }
    if (item.RecipientId) {
      out.Recipient = { Title: (state.users[item.RecipientId] || { Title: 'Someone' }).Title };
    }
    return out;
  }

  function queryItems(list, url) {
    let rows = list.items.slice();
    const filter = param(url, '$filter') || '';
    const author = /AuthorId eq (\d+)/.exec(filter);
    if (author) {
      rows = rows.filter((r) => r.AuthorId === Number(author[1]));
    }
    const pollKey = /PollKey eq '((?:[^']|'')*)'/.exec(filter);
    if (pollKey) {
      rows = rows.filter((r) => r.PollKey === pollKey[1].replace(/''/g, "'"));
    }
    if (/FSObjType eq 0/.test(filter)) {
      rows = rows.filter((r) => r.FSObjType === 0);
    }
    const since = /Created ge datetime'([^']+)'/.exec(filter);
    if (since) {
      rows = rows.filter((r) => r.Created >= since[1]);
    }
    const orderby = param(url, '$orderby') || '';
    if (/Id desc/.test(orderby)) {
      rows.sort((a, b) => b.Id - a.Id);
    } else {
      rows.sort((a, b) => a.Id - b.Id);
    }
    const top = Number(param(url, '$top') || 100);
    const skip = Number(param(url, 'mockskip') || 0);
    const page = rows.slice(skip, skip + top).map(expandUsers);
    const body = { value: page };
    if (skip + top < rows.length) {
      const base = url.replace(/[?&]mockskip=\d+/, '');
      body['@odata.nextLink'] = `${base}&mockskip=${skip + top}`;
    }
    return json(200, body);
  }

  function findList(url) {
    const byPath = quoted(url, "GetList('");
    if (byPath) {
      return state.lists[byPath.toLowerCase()];
    }
    const byTitle = quoted(url, "getbytitle('");
    if (byTitle) {
      return Object.keys(state.lists).map((k) => state.lists[k]).filter((l) => l.Title.toLowerCase() === byTitle.toLowerCase())[0];
    }
    const byId = /lists\(guid'([^']+)'\)/.exec(decodeURIComponent(url));
    if (byId) {
      return Object.keys(state.lists).map((k) => state.lists[k]).filter((l) => l.Id === byId[1])[0];
    }
    return undefined;
  }

  function searchQuery(url) {
    const kql = (param(url, 'querytext') || '').replace(/^'|'$/g, '').replace(/''/g, "'");
    const rowLimit = Number(param(url, 'rowlimit') || 10);
    const startRow = Number(param(url, 'startrow') || 0);
    const sort = (param(url, 'sortlist') || '').replace(/'/g, '');
    let rows = SEARCH_FILES.slice();
    const pathMatch = /Path:"([^"]+)\/\*"/.exec(kql);
    if (pathMatch) {
      rows = rows.filter((r) => r.Path.indexOf(pathMatch[1] + '/') === 0);
    }
    const before = /LastModifiedTime<(\d{4}-\d{2}-\d{2})/.exec(kql);
    if (before) {
      rows = rows.filter((r) => r.LastModifiedTime < before[1]);
    }
    const docId = /IndexDocId>(\d+)/.exec(kql);
    if (docId) {
      rows = rows.filter((r) => Number(r.DocId) > Number(docId[1]));
    }
    if (/CheckoutUserOWSUSER:/.test(kql)) {
      rows = options.state === 'empty' ? [] : rows.slice(3, 5);
    }
    if (/EditorOWSUSER:/.test(kql)) {
      rows = options.state === 'empty' ? [] : rows.filter((_, i) => i % 11 === 0);
    }
    if (/ViewableByExternalUsers:true/.test(kql)) {
      rows = options.state === 'empty' ? [] : rows.filter((_, i) => i % 6 === 0);
    }
    if (/Size:descending/.test(sort)) {
      rows.sort((a, b) => Number(b.Size) - Number(a.Size));
    } else if (/LastModifiedTime:ascending/.test(sort)) {
      rows.sort((a, b) => (a.LastModifiedTime < b.LastModifiedTime ? -1 : 1));
    } else if (/LastModifiedTime:descending/.test(sort)) {
      rows.sort((a, b) => (a.LastModifiedTime < b.LastModifiedTime ? 1 : -1));
    } else if (/DocId/.test(sort)) {
      rows.sort((a, b) => Number(a.DocId) - Number(b.DocId));
    }
    const page = rows.slice(startRow, startRow + rowLimit);
    const refiners = param(url, 'refiners')
      ? [
          { Name: 'FileType', Entries: [{ RefinementName: 'docx', RefinementCount: String(rows.length) }] },
          { Name: 'DisplayAuthor', Entries: [{ RefinementName: 'Priya Shah', RefinementCount: String(Math.floor(rows.length / 2)) }, { RefinementName: 'Tom Okafor', RefinementCount: String(Math.ceil(rows.length / 2)) }] }
        ]
      : [];
    // Answered in the OData 3 "verbose" shape, as the real search endpoint does.
    return json(200, {
      d: {
        query: {
          PrimaryQueryResult: {
            RelevantResults: {
              TotalRows: rows.length,
              Table: { Rows: { results: page.map((r) => ({ Cells: { results: Object.keys(r).map((k) => ({ Key: k, Value: r[k] })) } })) } }
            },
            RefinementResults: refiners.length ? { Refiners: { results: refiners.map((r) => ({ Name: r.Name, Entries: { results: r.Entries } })) } } : null
          }
        }
      }
    });
  }

  function roleAssignments() {
    return json(200, {
      value: [
        { Member: { Id: 3, Title: 'Team Owners', LoginName: 'Team Owners', PrincipalType: 8 }, RoleDefinitionBindings: [{ Name: 'Full Control' }] },
        { Member: { Id: 5, Title: 'Team Members', LoginName: 'Team Members', PrincipalType: 8 }, RoleDefinitionBindings: [{ Name: 'Edit' }, { Name: 'Limited Access' }] },
        { Member: state.users[10], RoleDefinitionBindings: [{ Name: 'Read' }] },
        { Member: { Id: 12, Title: 'Everyone except external users', LoginName: 'c:0-.f|rolemanager|spo-grid-all-users/x', PrincipalType: 4 }, RoleDefinitionBindings: [{ Name: 'Read' }] },
        { Member: { Id: 13, Title: 'Only limited', LoginName: 'x', PrincipalType: 1 }, RoleDefinitionBindings: [{ Name: 'Limited Access' }] }
      ]
    });
  }

  function folderRows(path) {
    const depth = path.split('/').length;
    return Array.from({ length: depth > 5 ? 0 : 3 }, (_, i) => ({
      Name: i === 2 ? 'Forms' : `Folder ${String.fromCharCode(65 + i)}`,
      ServerRelativeUrl: `${path}/${i === 2 ? 'Forms' : `Folder ${String.fromCharCode(65 + i)}`}`,
      StorageMetrics: { TotalSize: String((i + 1) * 3000000000), TotalFileStreamSize: String((i + 1) * 1200000000), TotalFileCount: String(40 * (i + 1)), LastModified: daysAgo(i * 20) }
    }));
  }

  function handleGet(url) {
    const u = decodeURIComponent(url);
    if (options.forbidden && new RegExp(options.forbidden).test(u)) {
      return denied();
    }
    if (/PeopleManager\/GetMyProperties/.test(u)) {
      if (options.noProfile) {
        return json(200, { AccountName: null, DisplayName: null });
      }
      return json(200, state.org[`i:0#.f|membership|${(state.users[me] || {}).Email || 'alex@contoso.com'}`] || { AccountName: null });
    }
    if (/PeopleManager\/GetPropertiesFor/.test(u)) {
      const account = /@v='(.*)'$/.exec(u)[1].replace(/''/g, "'").toLowerCase();
      return json(200, state.org[account] || { AccountName: null, DisplayName: null });
    }
    if (u.indexOf('/_api/search/query') >= 0) {
      return options.searchFails ? json(500, { error: { message: 'Search is unavailable.' } }) : searchQuery(url);
    }
    if (/\/_api\/web\/lists\?/.test(u)) {
      const filter = param(url, '$filter') || '';
      if (/BaseTemplate eq 119/.test(filter)) {
        return json(200, { value: [{ Id: 'lib-pages' }] });
      }
      if (/BaseTemplate eq 101/.test(filter)) {
        return json(200, { value: state.libraries.map((l) => ({ Id: l.Id, Title: l.Title, ItemCount: l.ItemCount, RootFolder: { ServerRelativeUrl: l.path } })) });
      }
      return json(200, {
        value: state.libraries.map((l) => ({ Id: l.Id, Title: l.Title, ItemCount: l.ItemCount, BaseTemplate: 101, HasUniqueRoleAssignments: !!l.HasUniqueRoleAssignments }))
          .concat([{ Id: 'l-tasks', Title: 'Tasks', ItemCount: 4, BaseTemplate: 171, HasUniqueRoleAssignments: true }])
      });
    }
    if (/\/_api\/web\/RootFolder\/Folders/.test(u)) {
      if (!owner) {
        return denied();
      }
      return json(200, {
        value: state.libraries.map((l, i) => ({
          Name: l.path.split('/').pop(), ServerRelativeUrl: l.path,
          StorageMetrics: { TotalSize: String((i + 1) * 9000000000), TotalFileStreamSize: String((i + 1) * 2500000000), TotalFileCount: String(l.ItemCount * 100), LastModified: daysAgo(i) }
        })).concat([{ Name: '_catalogs', ServerRelativeUrl: `${WEB}/_catalogs`, StorageMetrics: { TotalSize: '5', TotalFileStreamSize: '5', TotalFileCount: '1', LastModified: daysAgo(1) } }])
      });
    }
    const folderPath = quoted(url, "GetFolderByServerRelativePath(decodedurl='");
    if (folderPath) {
      if (/\/Files\?/.test(u)) {
        return json(200, {
          value: [
            { Name: 'Budget 2024.xlsx', ServerRelativeUrl: `${WEB}/Shared Documents/Budget 2024.xlsx`, Length: '2400000', TimeLastModified: daysAgo(3), UIVersionLabel: '15.0' },
            { Name: 'Tiny.txt', ServerRelativeUrl: `${folderPath}/Tiny.txt`, Length: '12', TimeLastModified: daysAgo(400), UIVersionLabel: '1.0' }
          ]
        });
      }
      return json(200, { value: folderRows(folderPath) });
    }
    const filePath = quoted(url, "GetFileByServerRelativePath(decodedurl='");
    if (filePath) {
      if (/\/Versions/.test(u)) {
        return json(200, { value: state.versions[filePath] || [] });
      }
      if (/RoleAssignments/.test(u)) {
        return owner ? roleAssignments() : denied();
      }
    }
    if (/\/_api\/(web|site)\/RecycleBin/.test(u)) {
      if (/\$orderby/.test(u) && options.noRecycleOrder) {
        return json(400, { error: { message: "The query option '$orderby' is not supported." } });
      }
      const top = Number(param(url, '$top') || 100);
      const visible = /\/_api\/site\//.test(u) ? state.recycle : state.recycle.filter((r) => r.ItemState === 1);
      return json(200, { value: visible.slice(0, top) });
    }
    if (/\/_api\/web\/RoleAssignments/.test(u)) {
      return owner ? roleAssignments() : denied();
    }
    if (/SiteGroups\/GetById\((\d+)\)\/Users/.test(u)) {
      const id = Number(/GetById\((\d+)\)/.exec(u)[1]);
      return json(200, { value: id === 3 ? [state.users[8]] : id === 5 ? [state.users[7], state.users[9], state.users[10]] : [] });
    }
    if (/AssociatedMemberGroup\/Users/.test(u)) {
      return json(200, { value: [state.users[7], state.users[8], state.users[9], state.users[11]] });
    }
    if (/\/_api\/site\/UserCustomActions/.test(u)) {
      return options.notSiteAdmin ? denied() : json(200, { value: state.customActions });
    }
    const list = findList(url);
    if (/GetList\(|getbytitle\(|lists\(guid/.test(u)) {
      if (!list) {
        return notFound();
      }
      if (/\/fields/.test(u)) {
        if (/Choices/.test(param(url, '$select') || '') && options.noChoicesSelect) {
          return json(400, { error: { message: "The field or property 'Choices' does not exist." } });
        }
        return json(200, { value: list.fieldDefs || list.fields.map((f) => ({ InternalName: f })) });
      }
      if (/\/items/.test(u)) {
        return queryItems(list, url);
      }
      return json(200, { Id: list.Id });
    }
    return json(500, { error: { message: `Mock SharePoint has no answer for GET ${u}` } });
  }

  function handlePost(url, init) {
    const u = decodeURIComponent(url);
    const headers = (init && init.headers) || {};
    const method = headers['X-HTTP-Method'] || 'POST';
    const body = init && init.body ? init.body : undefined;
    if (options.forbidden && new RegExp(options.forbidden).test(u)) {
      return denied();
    }
    if (/\/_api\/\$batch$/.test(u) && /^GET /m.test(body)) {
      const head = '--batchresponse_x\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\n';
      const reads = body.split(/\r\n/).filter((line) => line.indexOf('GET ') === 0).map((line) => line.substring(4, line.indexOf(' HTTP/1.1')));
      const out = reads.map((target, index) => {
        if (options.batchThrottleOnce && !state.batchThrottled && index === 1) {
          state.batchThrottled = true;
          return `${head}HTTP/1.1 429 Too Many Requests\r\nRetry-After: 2\r\n\r\n\r\n`;
        }
        const answer = handleGet(target);
        return `${head}HTTP/1.1 ${answer.status} X\r\nCONTENT-TYPE: application/json;odata=nometadata;charset=utf-8\r\n\r\n${answer.raw}\r\n`;
      });
      return json(200, `${out.join('')}--batchresponse_x--\r\n`);
    }
    if (/\/_api\/\$batch$/.test(u)) {
      const parts = body.split(/^PATCH /m).slice(1);
      const out = parts.map((part, index) => {
        const target = part.substring(0, part.indexOf(' HTTP/1.1'));
        const payload = JSON.parse(part.substring(part.indexOf('{'), part.lastIndexOf('}') + 1));
        const list = findList(target);
        const id = Number(/items\((\d+)\)/.exec(target)[1]);
        const item = list && list.items.filter((i) => i.Id === id)[0];
        const head = '--batchresponse_x\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\n';
        if (options.batchFailsItem === id) {
          return `${head}HTTP/1.1 400 Bad Request\r\nCONTENT-TYPE: application/json;odata=minimalmetadata;streaming=true;charset=utf-8\r\n\r\n{"error":{"code":"-1","message":"The file is locked for editing."}}\r\n`;
        }
        if (!item) {
          return `${head}HTTP/1.1 404 Not Found\r\n\r\n{"error":{"message":"Item does not exist."}}\r\n`;
        }
        Object.assign(item, payload);
        return `${head}HTTP/1.1 204 No Content\r\nCONTENT-TYPE: application/json;odata=minimalmetadata;streaming=true;charset=utf-8\r\n\r\n\r\n`;
      });
      return json(200, `${out.join('')}--batchresponse_x--\r\n`);
    }
    if (/clientPeoplePickerSearchUser/.test(u)) {
      const q = JSON.parse(body).queryParams.QueryString.toLowerCase();
      const hits = Object.keys(state.users).map((k) => state.users[k]).filter((p) => p.PrincipalType === 1 && p.Title.toLowerCase().indexOf(q) >= 0);
      return json(200, { value: JSON.stringify(hits.map((p) => ({ Key: p.LoginName, DisplayText: p.Title, EntityData: { Email: p.Email } }))) });
    }
    if (/\/_api\/web\/ensureuser/.test(u)) {
      const login = JSON.parse(body).logonName;
      const user = Object.keys(state.users).map((k) => state.users[k]).filter((p) => p.LoginName === login)[0];
      return user ? json(200, { Id: user.Id, Title: user.Title }) : json(400, { error: { message: 'The specified user could not be found.' } });
    }
    if (/RecycleBin\('([^']+)'\)\/restore/.test(u)) {
      const id = /RecycleBin\('([^']+)'\)/.exec(u)[1];
      if (id === 'rb-1') {
        return json(400, { error: { message: 'A file with this name already exists in the original location.' } });
      }
      state.recycle = state.recycle.filter((r) => r.Id !== id);
      return json(200, {});
    }
    if (/\/Versions\/RecycleByID\(vid=(\d+)\)/.test(u)) {
      const filePath = quoted(url, "GetFileByServerRelativePath(decodedurl='");
      const vid = Number(/vid=(\d+)/.exec(u)[1]);
      state.versions[filePath] = (state.versions[filePath] || []).filter((v) => v.ID !== vid);
      return json(200, {});
    }
    if (/\/copyTo\(/.test(u)) {
      const destination = quoted(url, "strNewUrl='");
      if (/exists/i.test(destination)) {
        return json(400, { error: { message: `A file with the name ${destination} already exists.` } });
      }
      return json(200, {});
    }
    if (/\/_api\/site\/UserCustomActions/.test(u)) {
      if (options.notSiteAdmin) {
        return denied();
      }
      if (method === 'DELETE') {
        state.customActions = [];
        return json(204);
      }
      const action = Object.assign({ Id: 'ca-1' }, JSON.parse(body));
      state.customActions.push(action);
      return json(201, action);
    }
    if (/\/_api\/web\/lists$/.test(u)) {
      if (!owner) {
        return denied();
      }
      const def = JSON.parse(body);
      const path = def.BaseTemplate === 101 ? `${WEB}/${def.Title}` : `${WEB}/Lists/${def.Title}`;
      const created = { Id: `new-${++state.nextId}`, Title: def.Title, path, items: [], fields: ['Title'], library: def.BaseTemplate === 101 };
      state.lists[path.toLowerCase()] = created;
      if (def.BaseTemplate === 101) {
        state.libraries.push({ Id: created.Id, Title: def.Title, ItemCount: 0, path, BaseTemplate: 101 });
      }
      return json(201, { Id: created.Id });
    }
    const list = findList(url);
    if (list) {
      if (/createfieldasxml/.test(u)) {
        const xml = JSON.parse(body).parameters.SchemaXml;
        list.fields.push(/ Name="([^"]+)"/.exec(xml)[1]);
        return json(200, {});
      }
      if (/\/fields\//.test(u)) {
        return json(204);
      }
      const itemMatch = /items\((\d+)\)/.exec(u);
      if (itemMatch) {
        const id = Number(itemMatch[1]);
        if (/recycle\(\)/.test(u)) {
          list.items = list.items.filter((i) => i.Id !== id);
          return json(200, {});
        }
        Object.assign(list.items.filter((i) => i.Id === id)[0] || {}, JSON.parse(body));
        return json(204);
      }
      if (/\/items$/.test(u)) {
        if (options.readOnly) {
          return denied();
        }
        const values = JSON.parse(body);
        for (const key of Object.keys(values)) {
          // Lists created through setup only accept the columns setup actually created.
          if (!list.seeded && key !== 'Title' && list.fields.indexOf(key.replace(/Id$/, '')) < 0 && list.fields.indexOf(key) < 0 && !list.library) {
            return json(400, { error: { message: `The property '${key}' does not exist on type 'SP.Data.ListItem'.` } });
          }
        }
        const item = Object.assign({ Id: ++state.nextId, AuthorId: me, Created: new Date().toISOString(), Modified: new Date().toISOString() }, values);
        list.items.push(item);
        return json(201, item);
      }
      if (method === 'MERGE') {
        Object.assign(list, JSON.parse(body));
        return json(204);
      }
    }
    if (/GetList\(|getbytitle\(/.test(u)) {
      return notFound();
    }
    return json(500, { error: { message: `Mock SharePoint has no answer for POST ${u}` } });
  }

  function send(method, url, init) {
    const entry = { method, url: decodeURIComponent(url), start: performance.now(), end: 0, status: 0, overlap: active > 0 };
    log.push(entry);
    active++;
    count++;
    return new Promise((resolve) => {
      setTimeout(() => {
        let result;
        if (options.throttleAt && options.throttleAt.indexOf(count) >= 0) {
          result = json(429, '', { 'Retry-After': '2' });
        } else {
          try {
            result = method === 'GET' ? handleGet(url) : handlePost(url, init);
          } catch (err) {
            result = json(500, { error: { message: `Mock SharePoint crashed: ${err && err.message}` } });
          }
        }
        active--;
        entry.end = performance.now();
        entry.status = result.status;
        resolve(result);
      }, latency);
    });
  }

  return {
    state,
    log,
    spHttpClient: {
      get: (url, config, init) => send('GET', url, init),
      post: (url, config, init) => send('POST', url, init)
    }
  };
}

module.exports = { createMock, ORIGIN, WEB };
