/**
 * The monthly Power Automate flow, generated as a workflow definition.
 *
 * It does what the web part's "Run now" does (services/AccessScan.ts, services/Sites.ts) and
 * writes the same file, using one connector only: SharePoint's "Send an HTTP request to
 * SharePoint" - a standard (not premium) action that runs as the flow owner's connection.
 * No app registration, no Microsoft Graph, no premium licence.
 *
 * Design rules, all checked by tests/flow.test.js against a simulator of the workflow language:
 *  - every expression is safe when `if` evaluates both branches (nothing relies on laziness);
 *  - integers stay inside 32 bits (the 64-bit permission mask is split before any arithmetic);
 *  - a request that fails is recorded and the run carries on: the action after it always
 *    runs and decides from the status code, so no branch is left failed or skipped;
 *  - every loop runs one item at a time, which is gentle on SharePoint and keeps the rows in
 *    the same order as the web part's;
 *  - no single expression comes near the service's 8,192 character limit.
 *
 * While it scans it appends text lines to one variable: report rows, and - prefixed "!" -
 * things it could not read. A person's row carries the raw permission mask between unit
 * separators; one Select at the end turns every mask into a level, so the long level
 * expression is evaluated in a single action rather than once per site.
 */

export const FLOW_NAME = 'Access Report - monthly';
export const PLACEHOLDER_URL = 'https://contoso.sharepoint.com/sites/your-site';
export const SHAREPOINT_API = '/providers/Microsoft.PowerApps/apis/shared_sharepointonline';
export const CONNECTION_KEY = 'shared_sharepointonline';

type Json = { [key: string]: unknown };
type Step = [string, Json, string[]?];

const ALWAYS = ['Succeeded', 'Failed', 'TimedOut'];
const CRLF = "decodeUriComponent('%0D%0A')";
/** Unit separator: cannot be typed into a title, so it can frame the raw mask inside a line. */
const U = "decodeUriComponent('%1F')";
const NONE = "json('[]')";
const JSON_ACCEPT = { Accept: 'application/json;odata=nometadata' };

/** Actions run in the order given; each runs after the one before it. */
function chain(steps: Step[]): Json {
  const out: Json = {};
  let prev: string | undefined;
  steps.forEach(([name, action, after]) => {
    out[name] = { ...action, runAfter: prev ? { [prev]: after || ['Succeeded'] } : {} };
    prev = name;
  });
  return out;
}

const compose = (inputs: unknown, description?: string): Json => (description ? { type: 'Compose', inputs, description } : { type: 'Compose', inputs });
const initVar = (name: string, type: string, value: unknown): Json => ({ type: 'InitializeVariable', inputs: { variables: [{ name, type, value }] } });
const setVar = (name: string, value: unknown): Json => ({ type: 'SetVariable', inputs: { name, value } });
const appendLines = (value: string): Json => ({ type: 'AppendToStringVariable', inputs: { name: 'Lines', value: '@' + value } });
const select = (from: string, map: unknown): Json => ({ type: 'Select', inputs: { from: '@' + from, select: map } });
const query = (from: string, where: string): Json => ({ type: 'Query', inputs: { from: '@' + from, where: '@' + where } });
const forEach = (items: string, actions: Json): Json => ({ type: 'Foreach', foreach: '@' + items, actions, runtimeConfiguration: { concurrency: { repetitions: 1 } } });

function http(method: 'GET' | 'POST', site: string, uri: string, extra: { headers?: Json; body?: string } = {}): Json {
  const parameters: Json = { dataset: '@' + site, 'parameters/method': method, 'parameters/uri': uri, 'parameters/headers': { ...JSON_ACCEPT, ...(extra.headers || {}) } };
  if (extra.body !== undefined) {
    parameters['parameters/body'] = extra.body;
  }
  return {
    type: 'OpenApiConnection',
    inputs: { host: { apiId: SHAREPOINT_API, connectionName: CONNECTION_KEY, operationId: 'HttpRequest' }, parameters, authentication: "@parameters('$authentication')" }
  };
}

// ---- expression building blocks --------------------------------------------------------

/** The request answered 200. */
const ok = (action: string): string => `equals(outputs('${action}')?['statusCode'],200)`;
/** Its body when it did, an empty object when it did not (a failed request's body may not be an object). */
const safeBody = (action: string): string => `if(${ok(action)},body('${action}'),json('{}'))`;
const rowsOfAction = (action: string): string => `coalesce(${safeBody(action)}?['value'],${NONE})`;
/** The lines a Select produced, each ended by CRLF. */
const joined = (action: string): string => `if(empty(body('${action}')),'',concat(join(body('${action}'),${CRLF}),${CRLF}))`;

/** A CSV cell exactly as services/reportCsv.ts writes it: formula guard, quotes doubled, quoted. */
export function cell(x: string): string {
  const guarded = `if(or(startsWith(${x},'='),startsWith(${x},'+'),startsWith(${x},'-'),startsWith(${x},'@'),startsWith(${x},decodeUriComponent('%09')),startsWith(${x},decodeUriComponent('%0D'))),concat('''',${x}),${x})`;
  return `concat('"',replace(${guarded},'"','""'),'"')`;
}

/** What a failed request said, on one line, at most 300 characters. */
function errorText(action: string): string {
  const raw = `string(coalesce(body('${action}'),''))`;
  const message = `first(split(first(split(last(split(${raw},'"message":"')),'","')),'\\r\\n'))`;
  return `replace(replace(substring(${message},0,min(length(${message}),300)),decodeUriComponent('%0D'),' '),decodeUriComponent('%0A'),' ')`;
}

/** An "unread" line (kept apart from report rows by its leading "!") when the request failed. */
function unreadLine(action: string, whoCell: string, stepCell: string): string {
  const line = `concat('!',${whoCell},',',${START},',',${stepCell},',"',string(coalesce(outputs('${action}')?['statusCode'],'')),'","',replace(${errorText(action)},'"','""'),'"',${CRLF})`;
  return `if(${ok(action)},'',${line})`;
}

const SITE = "items('For_each_site')";
/** "site url","site title" - the cells every line of a site starts with (after Who). */
const START = "outputs('Site_cells')?['start']";
/** "site title","site url" - the Name and URL cells of a row about the site itself. */
const SELF = "outputs('Site_cells')?['self']";

const low = (action: string): string => `string(coalesce(${safeBody(action)}?['Low'],'0'))`;
const high = (action: string): string => `string(coalesce(${safeBody(action)}?['High'],'0'))`;
/** Anything at all granted: the cheap test that keeps the many "nothing here" answers out of the variable. */
const someMask = (action: string): string => `and(${ok(action)},not(and(equals(${low(action)},'0'),equals(${high(action)},'0'))))`;
/** A person's row: the mask travels between unit separators until the last step names the level. */
const maskRow = (action: string, head: string, tail: string): string => `concat(${head},${U},${low(action)},${U},${high(action)},${U},${tail})`;
/** A finished row (group grants): same four-part shape with no mask. */
const plainRow = (line: string): string => `concat(${line},${U},${U},${U})`;

/**
 * services/permissions.ts permissionLevel() over the Low / High halves found in a line, or
 * the unit separator for "No access" (those lines are then dropped). The Low half can exceed
 * a 32-bit integer, so it is read as a (billions) and b (the last nine digits):
 * n = a * 10^9 + b, and bit k of n is ((a * q) + ((a * r + b) div 2^k)) mod 2 where
 * 10^9 = q * 2^k + r. Every intermediate stays below 2^31.
 */
export function levelExpression(lowText: string, highText: string): string {
  const padded = `concat('0000000000',${lowText})`;
  const a = `int(substring(${padded},0,sub(length(${padded}),9)))`;
  const b = `int(substring(${padded},sub(length(${padded}),9),9))`;
  const bit = (kind: number): string => {
    const power = Math.pow(2, kind - 1);
    const q = Math.floor(1e9 / power);
    const r = 1e9 - q * power;
    // Kinds 1-4 sit in the last digits alone (10^9 is a multiple of 2^9).
    return r === 0 ? `equals(mod(div(${b},${power}),2),1)` : `equals(mod(add(mul(${a},${q}),div(add(mul(${a},${r}),${b}),${power})),2),1)`;
  };
  return (
    `if(equals(${lowText},''),'',` +
    `if(and(equals(${lowText},'4294967295'),equals(${highText},'2147483647')),'Full control',` +
    `if(or(${bit(19)},${bit(26)}),'Design / manage',` +
    `if(and(${bit(2)},${bit(3)},${bit(4)}),'Edit / contribute',` +
    `if(${bit(1)},'Read',${U})))))`
  );
}

const HEADER_LINE = '"Who","Site URL","Site title","Object","Name","URL","Access","How","Own permissions"';
const UNREAD_HEADER_LINE = '"Who","Site URL","Site title","Step","Status","Error"';
const ASSIGNMENTS = 'roleassignments?$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings&$top=5000';

/** A grant's level names and how it reaches the group, for the loop over `grant` items. */
function grantParts(grant: string, namesAction: string): { level: string; how: string; none: string } {
  const how = `if(equals(${grant}?['Member']?['PrincipalType'],8),concat('SharePoint group: ',coalesce(${grant}?['Member']?['Title'],'')),'Direct')`;
  return { level: cell(`join(body('${namesAction}'),', ')`), how, none: `empty(body('${namesAction}'))` };
}

const NOT_LIMITED = "and(not(equals(item()?['RoleTypeKind'],1)),not(startsWith(toLower(coalesce(item()?['Name'],'')),'limited access')))";

// ---- the parts of the flow -------------------------------------------------------------

function listTheSites(): Json {
  const rows = `coalesce(body('Search_sites')?['PrimaryQueryResult']?['RelevantResults']?['Table']?['Rows'],${NONE})`;
  // A search row is a list of {Key, Value} cells: read one by its key.
  const cellOf = (row: string, key: string): string => `xpath(xml(json(concat('{"r":{"c":',string(coalesce(${row}?['Cells'],${NONE})),'}}'))),'string(/r/c[Key="${key}"]/Value)')`;
  const path = cellOf('item()', 'Path');
  const url = `substring(${path},0,sub(length(${path}),if(endsWith(${path},'/'),1,0)))`;
  const title = cellOf('item()', 'Title');
  const origin = "toLower(concat(outputs('Origin'),'/'))";
  const address = "toLower(concat(item()?['url'],'/'))";
  return {
    type: 'Until',
    expression: "@less(variables('PageSize'),500)",
    limit: { count: 400, timeout: 'PT6H' },
    actions: chain([
      // Set first, so a failed search ends the loop instead of repeating it.
      ['Assume_this_is_the_last_page', setVar('PageSize', 0)],
      [
        'Search_sites',
        http(
          'GET',
          "outputs('Home')",
          "_api/search/query?querytext='contentclass%3ASTS_Site%20IndexDocId%3E@{variables('LastDocId')}'&selectproperties='Title%2CPath%2CWebTemplate%2CDocId'&rowlimit=500&trimduplicates=false&sortlist='%5BDocId%5D%3Aascending'"
        )
      ],
      ['Site_rows', select(rows, { url: '@' + url, title: `@if(equals(${title},''),${url},${title})`, template: '@' + cellOf('item()', 'WebTemplate') })],
      [
        'Sites_in_scope',
        query(
          "body('Site_rows')",
          `and(not(equals(item()?['url'],'')),startsWith(${address},${origin}),not(startsWith(${address},toLower(concat(outputs('Origin'),'/personal/')))),not(startsWith(toUpper(item()?['template']),'SPSPERS')),contains(toLower(concat(item()?['title'],' ',item()?['url'])),outputs('Settings')?['siteFilter']))`
        )
      ],
      ['Sites_so_far', compose("@union(variables('Sites'),body('Sites_in_scope'))")],
      ['Remember_sites', setVar('Sites', "@outputs('Sites_so_far')")],
      ['Next_DocId', compose('@' + cellOf(`last(${rows})`, 'DocId'))],
      ['Remember_DocId', setVar('LastDocId', "@if(equals(outputs('Next_DocId'),''),'0',outputs('Next_DocId'))")],
      // No DocId on a full page would repeat the same page: stop instead.
      ['Remember_page_size', setVar('PageSize', `@if(equals(outputs('Next_DocId'),''),0,length(${rows}))`)]
    ])
  };
}

/** Reads the site's libraries once; every subject of the site then uses the two filtered lists. */
function librarySteps(): Step[] {
  const path = "coalesce(item()?['RootFolder']?['ServerRelativeUrl'],'')";
  const kind = "if(and(equals(item()?['BaseType'],1),not(equals(item()?['IsSystemList'],true))),if(equals(item()?['HasUniqueRoleAssignments'],true),'own','inherited'),'skip')";
  return [
    ['Get_libraries', http('GET', `${SITE}?['url']`, '_api/web/lists?$select=Id,Title,BaseType,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden%20eq%20false&$top=5000')],
    [
      'Libraries',
      select(rowsOfAction('Get_libraries'), {
        id: "@item()?['Id']",
        name: "@coalesce(item()?['Title'],'')",
        title: '@' + cell("coalesce(item()?['Title'],'')"),
        url: '@' + cell(`if(equals(${path},''),'',concat(outputs('Origin'),${path}))`),
        kind: '@' + kind
      }),
      ALWAYS
    ],
    ['Inherited_libraries', query("body('Libraries')", "equals(item()?['kind'],'inherited')")],
    ['Own_permission_libraries', query("body('Libraries')", "equals(item()?['kind'],'own')")],
    ['Libraries_unread', appendLines(unreadLine('Get_libraries', `'"(all)"'`, `'"Libraries"'`))]
  ];
}

/**
 * One person on one site. Built twice - with the library steps and (suffix "_site_only")
 * without - so a run that leaves libraries out spends two actions per person and site.
 */
function forEachPerson(libraries: boolean): Json {
  const x = libraries ? '' : '_site_only';
  const loop = `For_each_person${x}`;
  const person = `items('${loop}')`;
  const who = `${person}?['who']`;
  const check = `Check_person_on_site${x}`;
  const param = `'@{${person}?['param']}'`;
  const siteRow = `if(${someMask(check)},concat(${maskRow(check, `concat(${who},',',${START},',"Site",',${SELF},',"')`, `'","Effective permission",""'`)},${CRLF}),'')`;
  const steps: Step[] = [[check, http('GET', `${SITE}?['url']`, `_api/web/getusereffectivepermissions(@v)?@v=${param}`)]];
  if (!libraries) {
    steps.push([`Person_lines${x}`, appendLines(`concat(${siteRow},${unreadLine(check, who, `'"Site"'`)})`), ALWAYS]);
    return forEach("body('People')", chain(steps));
  }
  const libCheck = 'Check_person_on_library';
  const lib = "items('For_each_person_library')";
  steps.push(
    [
      'Person_library_lines',
      select(`if(${someMask(check)},body('Inherited_libraries'),${NONE})`, '@' + maskRow(check, `concat(${who},',',${START},',"Library",',item()?['title'],',',item()?['url'],',"')`, `'","Inherits from site","No"'`)),
      ALWAYS
    ],
    ['Person_lines', appendLines(`concat(${siteRow},${joined('Person_library_lines')},${unreadLine(check, who, `'"Site"'`)})`)],
    [
      'For_each_person_library',
      forEach(
        `if(${ok(check)},body('Own_permission_libraries'),${NONE})`,
        chain([
          [libCheck, http('GET', `${SITE}?['url']`, `_api/web/lists(guid'@{${lib}?['id']}')/getusereffectivepermissions(@v)?@v=${param}`)],
          [
            'Person_library_line',
            appendLines(
              `concat(if(${someMask(libCheck)},concat(${maskRow(
                libCheck,
                `concat(${who},',',${START},',"Library",',${lib}?['title'],',',${lib}?['url'],',"')`,
                `'","Effective permission (own permissions)","Yes"'`
              )},${CRLF}),''),${unreadLine(libCheck, who, cell(`concat('Library: ',${lib}?['name'])`))})`
            ),
            ALWAYS
          ]
        ])
      )
    ]
  );
  return forEach("body('People')", chain(steps));
}

/** One group (or other claim) on one site; built twice like forEachPerson. */
function forEachGroup(libraries: boolean): Json {
  const x = libraries ? '' : '_site_only';
  const loop = `For_each_group${x}`;
  const group = `items('${loop}')`;
  const claimLoop = `For_each_claim_found${x}`;
  const claim = `items('${claimLoop}')`;
  const find = `Find_group_on_site${x}`;
  const ids = `SharePoint_group_ids${x}`;
  const permissions = `Get_site_permissions${x}`;
  const grantLoop = `For_each_site_grant${x}`;
  const names = `Site_grant_level_names${x}`;
  const who = `if(and(equals(${group}?['m365'],true),endsWith(toLower(coalesce(${claim}?['LoginName'],'')),'_o')),${group}?['whoOwners'],${group}?['who'])`;
  const reaches = `or(equals(item()?['PrincipalId'],${claim}?['Id']),contains(body('${ids}'),item()?['PrincipalId']))`;
  const sg = grantParts(`items('${grantLoop}')`, names);
  const siteGrantRow = `if(${sg.none},'',concat(${plainRow(`concat(${who},',',${START},',"Site",',${SELF},',',${sg.level},',',${cell(sg.how)},',""')`)},${CRLF}))`;

  const grantSteps: Step[] = [
    [`Site_grant_levels${x}`, query(`coalesce(items('${grantLoop}')?['RoleDefinitionBindings'],${NONE})`, NOT_LIMITED)],
    [names, select(`body('Site_grant_levels${x}')`, "@coalesce(item()?['Name'],'')")]
  ];
  if (libraries) {
    grantSteps.push(
      [
        'Site_grant_library_lines',
        select(
          `if(${sg.none},${NONE},body('Inherited_libraries'))`,
          '@' + plainRow(`concat(${who},',',${START},',"Library",',item()?['title'],',',item()?['url'],',',${sg.level},',',${cell(`concat('Inherits from site - ',${sg.how})`)},',"No"')`)
        )
      ],
      ['Site_grant_lines', appendLines(`concat(${siteGrantRow},${joined('Site_grant_library_lines')})`)]
    );
  } else {
    grantSteps.push([`Site_grant_lines${x}`, appendLines(siteGrantRow)]);
  }

  const claimSteps: Step[] = [
    [ids, select(`coalesce(${claim}?['Groups'],${NONE})`, "@item()?['Id']")],
    [permissions, http('GET', `${SITE}?['url']`, `_api/web/${ASSIGNMENTS}`)],
    [`Site_grants${x}`, query(rowsOfAction(permissions), reaches), ALWAYS],
    [
      `Claim_lines${x}`,
      appendLines(
        `concat(if(equals(${claim}?['IsSiteAdmin'],true),concat(${plainRow(
          `concat(${who},',',${START},',"Site",',${SELF},',"Full control","Site collection administrator",""')`
        )},${CRLF}),''),${unreadLine(permissions, who, `'"Site permissions"'`)})`
      )
    ],
    [grantLoop, forEach(`body('Site_grants${x}')`, chain(grantSteps))]
  ];
  if (libraries) {
    const lib = "items('For_each_group_library')";
    const lg = grantParts("items('For_each_library_grant')", 'Library_grant_level_names');
    claimSteps.push([
      'For_each_group_library',
      forEach(
        `if(${ok(permissions)},body('Own_permission_libraries'),${NONE})`,
        chain([
          ['Get_library_permissions', http('GET', `${SITE}?['url']`, `_api/web/lists(guid'@{${lib}?['id']}')/${ASSIGNMENTS}`)],
          ['Library_grants', query(rowsOfAction('Get_library_permissions'), reaches), ALWAYS],
          [
            'For_each_library_grant',
            forEach(
              "body('Library_grants')",
              chain([
                ['Library_grant_levels', query(`coalesce(items('For_each_library_grant')?['RoleDefinitionBindings'],${NONE})`, NOT_LIMITED)],
                ['Library_grant_level_names', select("body('Library_grant_levels')", "@coalesce(item()?['Name'],'')")],
                [
                  'Library_grant_line',
                  appendLines(`if(${lg.none},'',concat(${plainRow(`concat(${who},',',${START},',"Library",',${lib}?['title'],',',${lib}?['url'],',',${lg.level},',',${cell(lg.how)},',"Yes"')`)},${CRLF}))`)
                ]
              ])
            )
          ],
          ['Library_permissions_unread', appendLines(unreadLine('Get_library_permissions', who, cell(`concat('Library: ',${lib}?['name'])`)))]
        ])
      )
    ]);
  }

  return forEach(
    "body('Groups')",
    chain([
      // Is the claim known to this site at all? Most sites end here: one request.
      [find, http('GET', `${SITE}?['url']`, `_api/web/siteusers?$filter=@{${group}?['filter']}&$select=Id,LoginName,IsSiteAdmin,Groups/Id,Groups/Title&$expand=Groups`)],
      [claimLoop, forEach(rowsOfAction(find), chain(claimSteps)), ALWAYS],
      [`Group_unread${x}`, appendLines(unreadLine(find, `${group}?['who']`, `'"Site"'`))]
    ])
  );
}

function forEachSite(): Json {
  const title = `${SITE}?['title']`;
  const url = `${SITE}?['url']`;
  const subjects = (libraries: boolean): Step[] => [
    [libraries ? 'For_each_person' : 'For_each_person_site_only', forEachPerson(libraries), libraries ? ALWAYS : undefined],
    [libraries ? 'For_each_group' : 'For_each_group_site_only', forEachGroup(libraries), ALWAYS]
  ];
  return forEach(
    "variables('Sites')",
    chain([
      ['Site_cells', compose({ start: `@concat(${cell(url)},',',${cell(title)})`, self: `@concat(${cell(title)},',',${cell(url)})` })],
      [
        'Libraries_wanted',
        {
          type: 'If',
          expression: { equals: ["@outputs('Settings')?['includeLibraries']", true] },
          actions: chain(librarySteps().concat(subjects(true))),
          else: { actions: chain(subjects(false)) }
        }
      ]
    ])
  );
}

export interface IFlowOptions {
  /** The site that hosts the Access Report web part (its lists and report folder). */
  homeUrl: string;
}

/** The workflow definition: trigger and actions, ready to be wrapped in either package format. */
export function buildDefinition(o: IFlowOptions): Json {
  const listAlias = (name: string): string => `@u='@{encodeUriComponent(concat(outputs('Web_path'),'/Lists/${name}'))}'`;
  const first = `first(coalesce(body('Get_settings')?['value'],${NONE}))`;
  const folder = `trim(coalesce(${first}?['ARReportFolder'],''))`;
  const home = "trim(outputs('Home_site_URL'))";
  const login = "trim(coalesce(item()?['ARLogin'],''))";
  const quoted = `replace(${login},'''','''''')`;
  const m365 = `contains(toLower(${login}),'|federateddirectoryclaimprovider|')`;
  const label = "coalesce(item()?['Title'],'')";
  const merge = { 'Content-Type': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' };
  const settingsItem = `_api/web/GetList(@u)/items(@{outputs('Settings')?['itemId']})?${listAlias('AccessReportSettings')}`;
  const now = "formatDateTime(utcNow(),'yyyy-MM-dd HH:mm')";
  const save = (name: string): string =>
    `_api/web/GetFolderByServerRelativeUrl(@f)/Files/add(url=@n,overwrite=true)?@f='@{encodeUriComponent(replace(outputs('Settings')?['folder'],'''',''''''))}'&@n='${name}'`;
  const file = (header: string, rowsAction: string): string => `@concat(decodeUriComponent('%EF%BB%BF'),'${header}',${CRLF},${joined(rowsAction)})`;
  const lineParts = `split(item(),${U})`;

  const actions = chain([
    ['Home_site_URL', compose(o.homeUrl.trim(), 'The site that hosts the Access Report web part. Its lists say who to report on and where the file goes.')],
    ['Home', compose(`@substring(${home},0,sub(length(${home}),if(endsWith(${home},'/'),1,0)))`)],
    ['Origin', compose("@concat('https://',uriHost(outputs('Home')))")],
    ['Web_path', compose("@if(equals(uriPath(outputs('Home')),'/'),'',uriPath(outputs('Home')))")],
    ['Initialize_Lines', initVar('Lines', 'string', '')],
    ['Initialize_Sites', initVar('Sites', 'array', [])],
    ['Initialize_LastDocId', initVar('LastDocId', 'string', '0')],
    ['Initialize_PageSize', initVar('PageSize', 'integer', 0)],
    ['Get_settings', http('GET', "outputs('Home')", `_api/web/GetList(@u)/items?$select=Id,ARIncludeLibraries,ARSiteFilter,ARReportFolder&$orderby=Id&$top=1&${listAlias('AccessReportSettings')}`)],
    [
      'Settings',
      compose({
        itemId: `@${first}?['Id']`,
        includeLibraries: `@not(equals(${first}?['ARIncludeLibraries'],false))`,
        siteFilter: `@toLower(trim(coalesce(${first}?['ARSiteFilter'],'')))`,
        folder: `@if(equals(${folder},''),concat(outputs('Web_path'),'/AccessReports'),${folder})`
      })
    ],
    ['Get_subjects', http('GET', "outputs('Home')", `_api/web/GetList(@u)/items?$select=Id,Title,ARLogin,ARActive&$orderby=Id&$top=5000&${listAlias('AccessReportSubjects')}`)],
    ['Active_subjects', query(`coalesce(body('Get_subjects')?['value'],${NONE})`, `and(equals(item()?['ARActive'],true),not(equals(${login},'')))`)],
    [
      'Subjects',
      select("body('Active_subjects')", {
        who: '@' + cell(label),
        whoOwners: '@' + cell(`concat(${label},' (owners)')`),
        person: `@startsWith(toLower(${login}),'i:')`,
        m365: '@' + m365,
        // The login as a URL parameter value, and as the filter that finds a group's claims in a site.
        param: `@encodeUriComponent(${quoted})`,
        filter: `@encodeUriComponent(if(${m365},concat('LoginName eq ''',${quoted},''' or LoginName eq ''',${quoted},'_o'''),concat('LoginName eq ''',${quoted},'''')))`
      })
    ],
    ['People', query("body('Subjects')", "equals(item()?['person'],true)")],
    ['Groups', query("body('Subjects')", "equals(item()?['person'],false)")],
    ['Mark_started', http('POST', "outputs('Home')", settingsItem, { headers: merge, body: `{"ARLastRun":"Started @{${now}} UTC"}` })],
    ['Until_all_sites_are_listed', listTheSites()],
    ['For_each_site', forEachSite()],
    // From here on the report is written even if the scan stopped part-way.
    ['All_lines', compose(`@split(variables('Lines'),${CRLF})`), ALWAYS],
    ['Report_lines', query("outputs('All_lines')", `startsWith(item(),'"')`)],
    ['Report_rows', select("body('Report_lines')", `@concat(${lineParts}[0],${levelExpression(`${lineParts}[1]`, `${lineParts}[2]`)},${lineParts}[3])`)],
    ['Rows_with_access', query("body('Report_rows')", `not(contains(item(),${U}))`)],
    ['Unread_lines', query("outputs('All_lines')", "startsWith(item(),'!')")],
    ['Unread_rows', select("body('Unread_lines')", '@substring(item(),1)')],
    ['File_stamp', compose("@formatDateTime(utcNow(),'yyyy-MM-dd-HHmm')")],
    ['Save_the_report', http('POST', "outputs('Home')", save("access-report-@{outputs('File_stamp')}.csv"), { body: file(HEADER_LINE, 'Rows_with_access') })],
    [
      'Anything_unread',
      {
        type: 'If',
        expression: { greater: ["@length(body('Unread_rows'))", 0] },
        actions: { Save_the_unread_list: { ...http('POST', "outputs('Home')", save("access-report-@{outputs('File_stamp')}-unread.csv"), { body: file(UNREAD_HEADER_LINE, 'Unread_rows') }), runAfter: {} } },
        else: { actions: {} }
      }
    ],
    [
      'Mark_finished',
      http('POST', "outputs('Home')", settingsItem, {
        headers: merge,
        body: `{"ARLastRun":"Finished @{${now}} UTC: @{length(variables('Sites'))} sites, @{length(body('Subjects'))} people and groups, @{length(body('Rows_with_access'))} rows, @{length(body('Unread_rows'))} not read. File: access-report-@{outputs('File_stamp')}.csv"}`
      })
    ]
  ]);
  // A leaf that is skipped when the scan stopped part-way, so the run shows as failed even
  // though the partial report was saved.
  actions.Scan_completed = { ...compose('The scan visited every site.'), runAfter: { For_each_site: ['Succeeded'] } };

  return {
    $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
    contentVersion: '1.0.0.0',
    parameters: { $connections: { defaultValue: {}, type: 'Object' }, $authentication: { defaultValue: {}, type: 'SecureObject' } },
    triggers: { On_the_first_of_every_month: { type: 'Recurrence', recurrence: { frequency: 'Month', interval: 1, startTime: '2026-01-01T03:00:00Z' } } },
    actions,
    outputs: {}
  };
}
