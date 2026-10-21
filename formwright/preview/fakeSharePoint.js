// An in-memory SharePoint site that answers the REST calls Formwright makes.
const WEB = 'https://contoso.sharepoint.com/sites/procurement';
const WEB_PATH = '/sites/procurement';

const LIST_IDS = {
  requests: '11111111-1111-1111-1111-111111111111',
  departments: '22222222-2222-2222-2222-222222222222',
  costCenters: '33333333-3333-3333-3333-333333333333',
  lines: '44444444-4444-4444-4444-444444444444',
  assets: '55555555-5555-5555-5555-555555555555'
};

const users = [
  { Id: 7, Title: 'Megan Bowen', Email: 'meganb@contoso.com', LoginName: 'i:0#.f|membership|meganb@contoso.com' },
  { Id: 8, Title: 'Alex Wilber', Email: 'alexw@contoso.com', LoginName: 'i:0#.f|membership|alexw@contoso.com' },
  { Id: 9, Title: 'Diego Siciliani', Email: 'diegos@contoso.com', LoginName: 'i:0#.f|membership|diegos@contoso.com' },
  { Id: 10, Title: 'Lynne Robbins', Email: 'lynner@contoso.com', LoginName: 'i:0#.f|membership|lynner@contoso.com' }
];

function f(InternalName, Title, TypeAsString, extra) {
  return Object.assign({ InternalName, EntityPropertyName: InternalName, Title, TypeAsString, Description: '', Required: false, ReadOnlyField: false, Hidden: false, DefaultValue: null }, extra || {});
}

const lists = {
  [LIST_IDS.requests]: {
    Title: 'Purchase Requests', BaseTemplate: 100, BaseType: 0, EnableAttachments: true, url: `${WEB_PATH}/Lists/PurchaseRequests`,
    fields: [
      f('ID', 'ID', 'Counter', { ReadOnlyField: true }),
      f('Title', 'Request title', 'Text', { Required: true, MaxLength: 255 }),
      f('RequestType', 'Request type', 'Choice', { Choices: ['Hardware', 'Software', 'Services', 'Travel'], DefaultValue: 'Hardware' }),
      f('Department', 'Department', 'Lookup', { LookupList: `{${LIST_IDS.departments}}`, LookupField: 'Title' }),
      f('CostCenter', 'Cost center', 'Lookup', { LookupList: `{${LIST_IDS.costCenters}}`, LookupField: 'Title' }),
      f('Quantity', 'Quantity', 'Number', { MinimumValue: 1, SchemaXml: '<Field Decimals="0" />' }),
      f('UnitPrice', 'Unit price', 'Currency', { MinimumValue: 0 }),
      f('Total', 'Total', 'Currency'),
      f('NeededBy', 'Needed by', 'DateTime', { DisplayFormat: 0 }),
      f('Urgent', 'Urgent', 'Boolean', { DefaultValue: '0' }),
      f('Justification', 'Business justification', 'Note', { RichText: false }),
      f('Tags', 'Tags', 'MultiChoice', { Choices: ['Budgeted', 'Recurring', 'Capex', 'Opex'] }),
      f('VendorSite', 'Vendor website', 'URL', { DisplayFormat: 0 }),
      f('Approver', 'Approver', 'User', { SelectionMode: 0 }),
      f('Status', 'Status', 'Choice', { Choices: ['Draft', 'Submitted', 'Approved', 'Rejected'], DefaultValue: 'Draft' }),
      f('ApproverComments', 'Approver comments', 'Note'),
      f('ApprovedOn', 'Approved on', 'DateTime', { DisplayFormat: 1 }),
      f('Attachments', 'Attachments', 'Attachments'),
      f('Author', 'Created By', 'User', { ReadOnlyField: true }),
      f('Created', 'Created', 'DateTime', { ReadOnlyField: true, DisplayFormat: 1 })
    ],
    items: [
      { Id: 1, Title: 'Laptops for new hires', RequestType: 'Hardware', DepartmentId: 2, CostCenterId: 4, Quantity: 3, UnitPrice: 1450, Total: 4350, NeededBy: '2026-11-15T00:00:00Z', Urgent: true, Justification: 'Three engineers start in November.', Tags: ['Budgeted', 'Capex'], VendorSite: { Url: 'https://www.contoso-hardware.com', Description: 'Contoso Hardware' }, ApproverId: 8, Status: 'Submitted', ApproverComments: null, ApprovedOn: null, AuthorId: 7, Created: '2026-10-12T09:30:00Z', AttachmentFiles: [{ FileName: 'quote.pdf', ServerRelativeUrl: `${WEB_PATH}/Lists/PurchaseRequests/Attachments/1/quote.pdf` }] }
    ]
  },
  [LIST_IDS.departments]: {
    Title: 'Departments', BaseTemplate: 100, BaseType: 0, EnableAttachments: false, url: `${WEB_PATH}/Lists/Departments`,
    fields: [f('ID', 'ID', 'Counter', { ReadOnlyField: true }), f('Title', 'Title', 'Text', { Required: true })],
    items: [{ Id: 1, Title: 'Finance' }, { Id: 2, Title: 'Engineering' }, { Id: 3, Title: 'Marketing' }, { Id: 4, Title: 'Operations' }]
  },
  [LIST_IDS.costCenters]: {
    Title: 'Cost Centers', BaseTemplate: 100, BaseType: 0, EnableAttachments: false, url: `${WEB_PATH}/Lists/CostCenters`,
    fields: [f('ID', 'ID', 'Counter', { ReadOnlyField: true }), f('Title', 'Title', 'Text', { Required: true }), f('Department', 'Department', 'Lookup', { LookupList: `{${LIST_IDS.departments}}`, LookupField: 'Title' })],
    items: [
      { Id: 1, Title: 'FIN-100 Accounting', DepartmentId: 1 }, { Id: 2, Title: 'FIN-200 Treasury', DepartmentId: 1 },
      { Id: 3, Title: 'ENG-100 Platform', DepartmentId: 2 }, { Id: 4, Title: 'ENG-200 Devices', DepartmentId: 2 },
      { Id: 5, Title: 'MKT-100 Brand', DepartmentId: 3 }, { Id: 6, Title: 'OPS-100 Facilities', DepartmentId: 4 }
    ]
  },
  [LIST_IDS.lines]: {
    Title: 'Request Lines', BaseTemplate: 100, BaseType: 0, EnableAttachments: false, url: `${WEB_PATH}/Lists/RequestLines`,
    fields: [
      f('ID', 'ID', 'Counter', { ReadOnlyField: true }), f('Title', 'Item', 'Text', { Required: true }),
      f('LineQty', 'Qty', 'Number'), f('LinePrice', 'Price', 'Currency'),
      f('Request', 'Request', 'Lookup', { LookupList: `{${LIST_IDS.requests}}`, LookupField: 'Title' })
    ],
    items: [
      { Id: 1, Title: 'ThinkPad X1 Carbon', LineQty: 3, LinePrice: 1300, RequestId: 1 },
      { Id: 2, Title: 'USB-C dock', LineQty: 3, LinePrice: 150, RequestId: 1 }
    ]
  },
  [LIST_IDS.assets]: { Title: 'Site Assets', BaseTemplate: 101, BaseType: 1, EnableAttachments: false, url: `${WEB_PATH}/SiteAssets`, fields: [], items: [] }
};

const files = {};
const fileVersions = {};
const folders = { [`${WEB_PATH}/SiteAssets`]: true };
const contentTypes = { [LIST_IDS.requests]: [{ StringId: '0x0100AA', Name: 'Item', Hidden: false }, { StringId: '0x012000BB', Name: 'Folder', Hidden: false }] };
export const log = [];

function response(status, body, headers) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return Promise.resolve({
    ok: status >= 200 && status < 300,
    status,
    statusText: String(status),
    headers: { get: (name) => (headers || {})[name] || null },
    json: () => Promise.resolve(JSON.parse(text)),
    text: () => Promise.resolve(text)
  });
}

function listFrom(url) {
  const m = /lists\(guid'([0-9a-f-]+)'\)/i.exec(url);
  return m ? { id: m[1].toLowerCase(), list: lists[m[1].toLowerCase()] } : undefined;
}

function expand(list, item) {
  const out = Object.assign({}, item);
  list.fields.forEach((fd) => {
    const p = fd.EntityPropertyName;
    if (fd.TypeAsString === 'Lookup' || fd.TypeAsString === 'LookupMulti') {
      const target = lists[fd.LookupList.replace(/[{}]/g, '').toLowerCase()];
      const id = item[`${p}Id`];
      const found = target && target.items.filter((i) => i.Id === id)[0];
      out[p] = found ? { Id: found.Id, Title: found.Title } : null;
    } else if (fd.TypeAsString === 'User') {
      const u = users.filter((x) => x.Id === item[`${p}Id`])[0];
      out[p] = u ? { Id: u.Id, Title: u.Title, EMail: u.Email, Name: u.LoginName } : null;
    }
  });
  out.ContentTypeId = '0x0100AA';
  return out;
}

function matches(item, filter) {
  if (!filter) return true;
  return filter.split(/\s+and\s+/i).every((raw) => {
    const clause = raw.replace(/^\(+|\)+$/g, '').trim();
    let m = /^substringof\('(.*)',(\w+)\)$/i.exec(clause);
    if (m) return String(item[m[2]] || '').toLowerCase().indexOf(m[1].replace(/''/g, "'").toLowerCase()) >= 0;
    m = /^(\w+) eq (\d+)$/i.exec(clause);
    if (m) return item[m[1]] === Number(m[2]);
    m = /^(\w+) eq '(.*)'$/i.exec(clause);
    if (m) return String(item[m[1]]) === m[2].replace(/''/g, "'");
    return true;
  });
}

function handleGet(url) {
  const decoded = decodeURIComponent(url);
  const filterMatch = /\$filter=([^&]*)/.exec(decoded);
  const filter = filterMatch ? filterMatch[1] : '';
  if (/_api\/web\/currentuser/.test(url)) {
    return response(200, { Id: 7, Title: 'Megan Bowen', Email: 'meganb@contoso.com', LoginName: users[0].LoginName, IsSiteAdmin: true, Groups: [{ Title: 'Procurement Owners' }, { Title: 'Finance Approvers' }] });
  }
  if (/_api\/web\/sitegroups/.test(url)) {
    return response(200, { value: ['Finance Approvers', 'Procurement Members', 'Procurement Owners', 'Procurement Visitors'].map((Title) => ({ Title })) });
  }
  if (/_api\/v2\.1\/termStore/.test(url)) return response(200, { value: [] });
  let m = /GetFileByServerRelativeUrl\('(.+)'\)\/\$value/.exec(decoded);
  if (m) {
    const path = m[1].replace(/''/g, "'");
    return files[path] ? response(200, files[path]) : response(404, { error: { message: 'File Not Found.' } });
  }
  m = /GetFileByServerRelativeUrl\('(.+)'\)\?\$select=ETag/.exec(decoded);
  if (m) {
    const path = m[1].replace(/''/g, "'");
    return files[path] ? response(200, { ETag: `"{F00D},${fileVersions[path] || 1}"` }) : response(404, { error: { message: 'File Not Found.' } });
  }
  m = /GetFolderByServerRelativeUrl\('(.+)'\)\?\$select=Exists/.exec(decoded);
  if (m) return folders[m[1]] ? response(200, { Exists: true }) : response(404, { error: { message: 'File Not Found.' } });
  const target = listFrom(url);
  if (target && target.list) {
    const list = target.list;
    if (/\/fields\?/.test(url)) return response(200, { value: list.fields });
    if (/\/EffectiveBasePermissions/.test(url)) return response(200, { High: 2147483647, Low: 4294967295 });
    if (/\/contenttypes/.test(url)) return response(200, { value: (contentTypes[target.id] || [{ StringId: '0x01', Name: 'Item', Hidden: false }]).map((c) => Object.assign({}, c)) });
    m = /\/items\((\d+)\)/.exec(url);
    if (m) {
      const item = list.items.filter((i) => i.Id === Number(m[1]))[0];
      if (!item) return response(404, { error: { message: 'Item does not exist.' } });
      return response(200, Object.assign({ '@odata.etag': `"${item.__v || 1}"` }, expand(list, item)));
    }
    if (/\/items\?/.test(url)) return response(200, { value: list.items.filter((i) => matches(i, filter)).map((i) => expand(list, i)) });
    return response(200, { Id: target.id, Title: list.Title, BaseTemplate: list.BaseTemplate, ItemCount: list.items.length, BaseType: list.BaseType, EnableAttachments: list.EnableAttachments, RootFolder: { ServerRelativeUrl: list.url } });
  }
  if (/_api\/web\/lists\?/.test(url)) {
    const all = Object.keys(lists).map((id) => ({ Id: id, Title: lists[id].Title, BaseTemplate: lists[id].BaseTemplate, ItemCount: lists[id].items.length, BaseType: lists[id].BaseType, RootFolder: { ServerRelativeUrl: lists[id].url } }));
    return response(200, { value: /BaseTemplate eq 101/.test(decoded) ? all.filter((l) => l.BaseTemplate === 101) : all.filter((l) => l.BaseTemplate !== 101) });
  }
  return response(404, { error: { message: `Fake SharePoint has no route for GET ${url}` } });
}

function handlePost(url, options) {
  const decoded = decodeURIComponent(url);
  const headers = (options && options.headers) || {};
  const method = headers['X-HTTP-Method'];
  const body = options && typeof options.body === 'string' && options.body ? JSON.parse(options.body) : undefined;
  log.push({ url: decoded, method: method || 'POST', body: options && typeof options.body === 'string' ? body : '(binary)' });
  let m = /GetFolderByServerRelativeUrl\('(.+)'\)\/Files\/add\(url='(.+)',overwrite=true\)/.exec(decoded);
  if (m) {
    files[`${m[1]}/${m[2]}`] = options.body;
    fileVersions[`${m[1]}/${m[2]}`] = (fileVersions[`${m[1]}/${m[2]}`] || 1) + 1;
    return response(200, {});
  }
  m = /GetFolderByServerRelativeUrl\('(.+)'\)\/Folders\/add\('(.+)'\)/.exec(decoded);
  if (m) {
    folders[`${m[1]}/${m[2]}`] = true;
    return response(200, {});
  }
  if (/clientPeoplePickerSearchUser/.test(url)) {
    const q = body.queryParams.QueryString.toLowerCase();
    const found = users.filter((u) => u.Title.toLowerCase().indexOf(q) >= 0 || u.Email.indexOf(q) >= 0);
    return response(200, { value: JSON.stringify(found.map((u) => ({ Key: u.LoginName, DisplayText: u.Title, EntityType: 'User', EntityData: { Email: u.Email, Title: 'Contoso' } }))) });
  }
  if (/_api\/web\/ensureuser/.test(url)) {
    const key = body.logonName.toLowerCase();
    const u = users.filter((x) => x.LoginName.toLowerCase() === key || x.Email.toLowerCase() === key)[0];
    return u ? response(200, u) : response(404, { error: { message: 'The specified user could not be found.' } });
  }
  const target = listFrom(url);
  if (target && target.list) {
    const list = target.list;
    m = /\/contenttypes\('(.+)'\)/.exec(decoded);
    if (m) {
      const ct = (contentTypes[target.id] || []).filter((c) => c.StringId === m[1])[0];
      if (ct) Object.assign(ct, body);
      return response(204);
    }
    if (/ValidateUpdateListItem/.test(url)) return response(200, { value: [] });
    m = /\/items\((\d+)\)\/AttachmentFiles\/add\(FileName='(.+)'\)/.exec(decoded);
    if (m) return response(200, {});
    m = /\/items\((\d+)\)$/.exec(decoded);
    if (m) {
      const idx = list.items.map((i) => i.Id).indexOf(Number(m[1]));
      if (method === 'DELETE') {
        list.items.splice(idx, 1);
        return response(204);
      }
      const item = list.items[idx];
      const etag = headers['IF-MATCH'];
      if (etag && etag !== '*' && etag !== `"${item.__v || 1}"`) return response(412, { error: { message: 'The version of the item does not match.' } });
      Object.assign(item, body, { __v: (item.__v || 1) + 1 });
      return response(204);
    }
    if (/\/items$/.test(decoded)) {
      const id = list.items.reduce((max, i) => Math.max(max, i.Id), 0) + 1;
      const item = Object.assign({ Id: id, AuthorId: 7, Created: new Date().toISOString() }, body);
      list.items.push(item);
      return response(201, { Id: id });
    }
  }
  return response(404, { error: { message: `Fake SharePoint has no route for POST ${url}` } });
}

export const fakeContext = {
  spHttpClient: {
    get: (url) => handleGet(url),
    post: (url, _config, options) => handlePost(url, options)
  },
  pageContext: {
    web: { absoluteUrl: WEB, serverRelativeUrl: WEB_PATH },
    site: { absoluteUrl: WEB },
    cultureInfo: { currentUICultureName: 'en-US' }
  }
};

export const fakeData = { lists, files, fileVersions, LIST_IDS, users };
