// Tenants shared by the parity tests: built fresh for each side (web part / flow), since a
// run writes files and settings into them.
const { FakeTenant, ORIGIN, FULL, READ, EDIT, DESIGN, LIMITED, NONE } = require('./fakeTenant');

const ANN = 'i:0#.f|membership|ann@contoso.com';
const GUEST = "i:0#.f|membership|o'brien_fabrikam.com#ext#@contoso.onmicrosoft.com";
const GHOST = 'i:0#.f|membership|ghost@contoso.com';
const AUDITORS = 'c:0t.c|tenant|11111111-1111-1111-1111-111111111111';
const MARKETING = 'c:0o.c|federateddirectoryclaimprovider|22222222-2222-2222-2222-222222222222';
// Not a real claim: characters that break a URL unless the flow encodes them.
const ODD = 'c:0-.f|rolemanager|s&p#1+x';
const EVERYONE = 'c:0-.f|rolemanager|spo-grid-all-users/33333333-3333-3333-3333-333333333333';

const role = (Name, RoleTypeKind = 0) => ({ Name, RoleTypeKind });
const grant = (PrincipalId, title, type, ...roles) => ({ PrincipalId, Member: { Title: title, PrincipalType: type }, RoleDefinitionBindings: roles });
const lib = (Id, Title, path, extra = {}) => Object.assign({ Id, Title, BaseType: 1, HasUniqueRoleAssignments: false, IsSystemList: false, RootFolder: path === null ? undefined : { ServerRelativeUrl: path } }, extra);

const SUBJECTS = [
  { label: 'Ann Example <ann@contoso.com>', login: ANN },
  { label: 'O\'Brien, "Pat" (guest)', login: GUEST },
  { label: 'Finance Auditors', login: AUDITORS },
  { label: '=Marketing team', login: MARKETING },
  { label: 'Everyone except external users', login: EVERYONE },
  { label: 'Odd claim', login: ODD },
  { label: 'Paused person', login: 'i:0#.f|membership|paused@contoso.com', active: false },
  { label: 'No login', login: '   ' },
  { label: 'Ghost', login: GHOST }
];

function richSites() {
  return [
    {
      url: ORIGIN + '/sites/it',
      title: 'IT',
      users: { [ANN]: FULL },
      siteUsers: [
        { Id: 20, LoginName: AUDITORS, Groups: [{ Id: 5, Title: 'IT Members' }] },
        { Id: 30, LoginName: MARKETING, Groups: [{ Id: 5, Title: 'IT Members' }, { Id: 7, Title: 'IT Visitors' }] },
        { Id: 31, LoginName: MARKETING + '_o', IsSiteAdmin: true, Groups: [{ Id: 3, Title: 'IT Owners' }] },
        { Id: 40, LoginName: EVERYONE.toUpperCase(), Groups: [] },
        { Id: 50, LoginName: ODD, Groups: [] }
      ],
      assignments: [
        grant(3, 'IT Owners', 8, role('Full Control', 5)),
        grant(5, 'IT Members', 8, role('Edit', 6), role('Limited Access', 1), role('Custom "Approve"')),
        grant(7, 'IT Visitors', 8, role('Beperkte toegang', 1)),
        grant(20, 'Finance Auditors', 4, role('Limited Access Something')),
        grant(40, 'Everyone except external users', 4, role('Read', 2)),
        { PrincipalId: 30, RoleDefinitionBindings: [role('=Odd level')] },
        grant(50, 'Odd claim', 4, role('Read', 2)),
        grant(99, 'Someone else', 1, role('Full Control', 5))
      ],
      lists: [
        lib('aaaaaaaa-0000-0000-0000-000000000001', 'Documents', '/sites/it/Shared Documents'),
        lib('aaaaaaaa-0000-0000-0000-000000000002', 'Secret "Plans", v2', '/sites/it/Secret Plans', {
          HasUniqueRoleAssignments: true,
          users: { [ANN]: FULL, [GUEST]: READ },
          assignments: [grant(20, 'Finance Auditors', 4, role('Read', 2), role('Contribute', 3)), grant(5, 'IT Members', 8, role('Limited Access', 1)), grant(31, 'Marketing Owners', 4, role('Design', 4))]
        }),
        lib('aaaaaaaa-0000-0000-0000-000000000003', 'Site Assets', '/sites/it/SiteAssets', { IsSystemList: true }),
        lib('aaaaaaaa-0000-0000-0000-000000000004', 'Tasks', '/sites/it/Lists/Tasks', { BaseType: 0 }),
        lib('aaaaaaaa-0000-0000-0000-000000000005', 'Hidden library', '/sites/it/Hidden', { Hidden: true }),
        lib('aaaaaaaa-0000-0000-0000-000000000006', null, null),
        lib('aaaaaaaa-0000-0000-0000-000000000007', 'Résumé – 日本語', '/sites/it/Resume', { HasUniqueRoleAssignments: true, users: {}, assignments: [] })
      ]
    },
    {
      url: ORIGIN + '/sites/fin',
      title: 'Finance, =SUM(1)',
      users: { [ANN]: READ, [GUEST]: LIMITED },
      siteUsers: [{ Id: 9, LoginName: AUDITORS, Groups: [] }],
      assignments: [grant(9, 'Finance Auditors', 4, role('Read', 2))],
      lists: [
        lib('bbbbbbbb-0000-0000-0000-000000000001', '+Budget', '/sites/fin/Budget'),
        lib('bbbbbbbb-0000-0000-0000-000000000002', 'Invoices', '/sites/fin/Invoices'),
        lib('bbbbbbbb-0000-0000-0000-000000000003', 'Payroll', '/sites/fin/Payroll', { HasUniqueRoleAssignments: true, users: { [GUEST]: EDIT, [ANN]: NONE }, assignments: [] }),
        lib('bbbbbbbb-0000-0000-0000-000000000004', 'Locked "vault"', '/sites/fin/Vault', { HasUniqueRoleAssignments: true, fail: 404 })
      ]
    },
    { url: ORIGIN + '/sites/hr', title: 'HR', deny: 403 },
    { url: ORIGIN + '/sites/legal', title: 'Legal', listsFail: 500, users: { [ANN]: DESIGN }, siteUsers: [], assignments: [] },
    { url: ORIGIN + '/sites/ops', title: 'Ops', assignmentsFail: 500, users: {}, siteUsers: [{ Id: 4, LoginName: AUDITORS, IsSiteAdmin: true, Groups: [] }], lists: [lib('cccccccc-0000-0000-0000-000000000001', 'Runbooks', '/sites/ops/Runbooks', { HasUniqueRoleAssignments: true })] },
    { url: ORIGIN, searchPath: ORIGIN + '/', title: 'Home', users: { [ANN]: EDIT, [GUEST]: READ }, siteUsers: [], assignments: [], lists: [] },
    { url: ORIGIN + '/personal/ann_contoso_com', title: 'Ann', template: 'SPSPERS' },
    { url: ORIGIN + '/sites/odd', title: 'Odd', template: 'spspers#10' },
    { url: ORIGIN + '/sites/empty', users: {}, siteUsers: [], assignments: [], lists: [] },
    { url: ORIGIN + '/sites/broken', title: 'Broken', broken: { status: 500, body: '<html>Server Error</html>', raw: '<html>Server Error</html>' } },
    { url: ORIGIN + '/sites/offline', title: 'Offline', broken: { transport: true } }
  ];
}

const OTHER_HOSTS = [
  { url: 'https://contoso-my.sharepoint.com/personal/bob_contoso_com', title: 'Bob', template: 'SPSPERS' },
  { url: 'https://fabrikam.sharepoint.com/sites/it', title: 'Elsewhere' },
  { url: ORIGIN + 'x/sites/lookalike', title: 'Look-alike host' }
];

function rich(options = {}) {
  const t = new FakeTenant({ home: options.home, sites: richSites() }).provisioned(Object.assign({ subjects: SUBJECTS }, options));
  t.extraSearchHits = OTHER_HOSTS;
  t.unknownUsers = [GHOST];
  return t;
}

/** Many plain sites, to walk more than one page of search results. */
function large(count) {
  const sites = [];
  for (let i = 0; i < count; i++) {
    sites.push({ url: `${ORIGIN}/sites/s${i}`, title: `Site ${i}`, users: i % 3 === 0 ? { [ANN]: i % 2 ? READ : EDIT } : {}, siteUsers: i % 50 === 0 ? [{ Id: 11, LoginName: AUDITORS, Groups: [] }] : [], assignments: [grant(11, 'Finance Auditors', 4, role('Read', 2))] });
  }
  return new FakeTenant({ sites }).provisioned({ subjects: [SUBJECTS[0], SUBJECTS[2]], includeLibraries: false });
}

module.exports = { rich, large, richSites, SUBJECTS, ANN, GUEST, GHOST, AUDITORS, MARKETING, EVERYONE, ORIGIN };
