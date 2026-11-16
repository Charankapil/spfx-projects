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

/** The web being scanned: a site collection's top-level site or a subsite. */
const SITE = "items('For_each_web')";
/** "site url","site title" - the cells every line of a site starts with (after Who). */
const START = "outputs('Site_cells')?['start']";
/** "site title","site url" - the Name and URL cells of a row about the site itself. */
const SELF = "outputs('Site_cells')?['self']";
/** "Site" or "Subsite" - the Object cell of a row about the web itself. */
const KIND = "outputs('Site_cells')?['kind']";
/** "" (top-level site), "Yes" or "No" (subsite with / without its own permissions). */
const OWN_WEB = "outputs('Site_cells')?['own']";
/** ,"<site collection url>" - the last cell of every row. */
const END = "outputs('Site_cells')?['end']";

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

const HEADER_LINE = '"Who","Site URL","Site title","Object","Name","URL","Access","How","Own permissions","Site collection"';
const UNREAD_HEADER_LINE = '"Who","Site URL","Site title","Step","Status","Error"';
const ASSIGNMENTS = 'roleassignments?$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings&$top=5000';

/** A grant's level names and how it reaches the group, for the loop over `grant` items. */
function grantParts(grant: string, namesAction: string): { level: string; how: string; none: string } {
  const how = `if(equals(${grant}?['Member']?['PrincipalType'],8),concat('SharePoint group: ',coalesce(${grant}?['Member']?['Title'],'')),'Direct')`;
  return { level: cell(`join(body('${namesAction}'),', ')`), how, none: `empty(body('${namesAction}'))` };
}

const NOT_LIMITED = "and(not(equals(item()?['RoleTypeKind'],1)),not(startsWith(toLower(coalesce(item()?['Name'],'')),'limited access')))";

// ---- the parts of the flow -------------------------------------------------------------

/**
 * Orderings for paging the site list, most stable first. Created never changes, so new sites
 * land at the end and no page shifts; LastModifiedTime is what most tenants accept; no order
 * at all is the last resort (pages can then overlap - the union drops repeats).
 * services/Sites.ts SITE_ORDERS is the same list.
 */
export const SITE_ORDERS = ['Created:ascending', 'LastModifiedTime:ascending', ''];

const sortParam = (o: string): string => (o ? `&sortlist='${o}'` : '');

function siteProbe(order: string): Json {
  return http('GET', "outputs('Home')", `_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Path'&rowlimit=1&trimduplicates=false${sortParam(order)}`);
}

/** The sortlist parameter of the first probe that worked, or 'refused'. */
function siteOrderExpression(): string {
  return SITE_ORDERS.reduceRight((rest, o, i) => `if(${ok(`Probe_site_order_${i + 1}`)},'${sortParam(o).replace(/'/g, "''")}',${rest})`, "'refused'");
}

/** What each probe said, JSON-safe, for the "Failed" note. */
function probeErrors(): string {
  const parts = SITE_ORDERS.map((o, i) => {
    const a = `Probe_site_order_${i + 1}`;
    return `concat('${o || 'no order'}: ',string(coalesce(outputs('${a}')?['statusCode'],'')),' ',${errorText(a)})`;
  });
  return `replace(replace(concat(${parts.join(",'; ',")}),'\\','\\\\'),'"','\\"')`;
}

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
    // 100 pages of 500: search serves a result set up to 50,000 rows this way.
    limit: { count: 100, timeout: 'PT6H' },
    actions: chain([
      // Set first, so a failed search ends the loop instead of repeating it.
      ['Assume_this_is_the_last_page', setVar('PageSize', 0)],
      [
        'Search_sites',
        // The plain site query, walked with startrow in the order the probes found accepted.
        http(
          'GET',
          "outputs('Home')",
          "_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Title,Path,WebTemplate'&rowlimit=500&startrow=@{variables('StartRow')}&trimduplicates=false@{outputs('Site_order')}"
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
      ['Next_start_row', compose("@add(variables('StartRow'),500)")],
      ['Remember_start_row', setVar('StartRow', "@outputs('Next_start_row')")],
      ['Remember_page_size', setVar('PageSize', `@length(${rows})`)]
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
  const siteRow = `if(${someMask(check)},concat(${maskRow(check, `concat(${who},',',${START},',',${KIND},',',${SELF},',"')`, `concat('","Effective permission",',${OWN_WEB},${END})`)},${CRLF}),'')`;
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
      select(`if(${someMask(check)},body('Inherited_libraries'),${NONE})`, '@' + maskRow(check, `concat(${who},',',${START},',"Library",',item()?['title'],',',item()?['url'],',"')`, `concat('","Inherits from site","No"',${END})`)),
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
                `concat('","Effective permission (own permissions)","Yes"',${END})`
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
  const siteGrantRow = `if(${sg.none},'',concat(${plainRow(`concat(${who},',',${START},',',${KIND},',',${SELF},',',${sg.level},',',${cell(sg.how)},',',${OWN_WEB},${END})`)},${CRLF}))`;

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
          '@' + plainRow(`concat(${who},',',${START},',"Library",',item()?['title'],',',item()?['url'],',',${sg.level},',',${cell(`concat('Inherits from site - ',${sg.how})`)},',"No"',${END})`)
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
          `concat(${who},',',${START},',',${KIND},',',${SELF},',"Full control","Site collection administrator",',${OWN_WEB},${END})`
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
                  appendLines(`if(${lg.none},'',concat(${plainRow(`concat(${who},',',${START},',"Library",',${lib}?['title'],',',${lib}?['url'],',',${lg.level},',',${cell(lg.how)},',"Yes"',${END})`)},${CRLF}))`)
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

/**
 * Every web of the current site collection into variables('Webs'): the top-level site, then
 * its subsites level by level (breadth first, one request per web - the order of
 * services/AccessScan.ts webs()). The top-level site's subsites are listed outside the loop,
 * so a site collection without subsites - most of them - costs six actions; the queue loop
 * only runs where there are subsites.
 */
function listTheWebs(): Step[] {
  const collection = "items('For_each_site')";
  /** The subsites a Get_... action returned, as web records. */
  const subsites = (action: string): Json => {
    const u = "coalesce(item()?['Url'],'')";
    const trimmed = `if(endsWith(${u},'/'),substring(${u},0,sub(length(${u}),1)),${u})`;
    return select(rowsOfAction(action), {
      url: '@' + trimmed,
      title: `@if(equals(coalesce(item()?['Title'],''),''),${trimmed},item()?['Title'])`,
      kind: 'Subsite',
      own: "@if(equals(item()?['HasUniqueRoleAssignments'],true),'Yes','No')",
      collection: `@${collection}?['url']`
    });
  };
  /** The "could not list the subsites" line for `web`, when `action` failed. */
  const unread = (action: string, web: string): Json =>
    appendLines(
      `if(${ok(action)},'',concat('!','"(all)"',',',${cell(`${web}?['url']`)},',',${cell(`${web}?['title']`)},',"Subsites","',string(coalesce(outputs('${action}')?['statusCode'],'')),'","',replace(${errorText(action)},'"','""'),'"',${CRLF}))`
    );
  const WEBS = '_api/web/webs?$select=Title,Url,HasUniqueRoleAssignments';
  const current = "first(variables('Queue'))";
  return [
    ['Top_level_site', compose({ url: `@${collection}?['url']`, title: `@${collection}?['title']`, kind: 'Site', own: '', collection: `@${collection}?['url']` })],
    ['Get_top_level_subsites', http('GET', `${collection}?['url']`, WEBS)],
    ['Top_level_subsites', subsites('Get_top_level_subsites'), ALWAYS],
    ['Top_level_subsites_unread', unread('Get_top_level_subsites', "outputs('Top_level_site')")],
    ['Start_the_webs', setVar('Webs', "@union(createArray(outputs('Top_level_site')),body('Top_level_subsites'))")],
    [
      'Any_subsites',
      {
        type: 'If',
        expression: { greater: ["@length(body('Top_level_subsites'))", 0] },
        actions: chain([
          ['Start_the_queue', setVar('Queue', "@body('Top_level_subsites')")],
          [
            'Until_all_subsites_are_listed',
            {
              type: 'Until',
              expression: "@equals(length(variables('Queue')),0)",
              limit: { count: 5000, timeout: 'PT6H' },
              actions: chain([
                ['Get_subsites', http('GET', `${current}?['url']`, WEBS)],
                ['Subsites', subsites('Get_subsites'), ALWAYS],
                ['Subsites_unread', unread('Get_subsites', current)],
                // The web just listed leaves the queue whatever happened, so a failure cannot repeat for ever.
                ['Next_queue', compose("@union(skip(variables('Queue'),1),body('Subsites'))")],
                ['Remember_the_queue', setVar('Queue', "@outputs('Next_queue')")],
                ['Webs_so_far', compose("@union(variables('Webs'),body('Subsites'))")],
                ['Remember_the_webs', setVar('Webs', "@outputs('Webs_so_far')")]
              ])
            }
          ]
        ]),
        else: { actions: {} }
      }
    ]
  ];
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
    chain(
      listTheWebs().concat([
        [
          'For_each_web',
          forEach(
            "variables('Webs')",
            chain([
              [
                'Site_cells',
                compose({
                  start: `@concat(${cell(url)},',',${cell(title)})`,
                  self: `@concat(${cell(title)},',',${cell(url)})`,
                  kind: '@' + cell(`${SITE}?['kind']`),
                  own: '@' + cell(`${SITE}?['own']`),
                  end: `@concat(',',${cell(`${SITE}?['collection']`)})`
                })
              ],
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
          ),
          ALWAYS
        ]
      ])
    )
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
    ['Initialize_StartRow', initVar('StartRow', 'integer', 0)],
    ['Initialize_Webs', initVar('Webs', 'array', [])],
    ['Initialize_Queue', initVar('Queue', 'array', [])],
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
    // Search on some tenants refuses some sort orders ("We didn't understand your search terms").
    // One-row probes find the first ordering this tenant accepts; services/Sites.ts tries the same list.
    ...SITE_ORDERS.map((o, i): Step => [`Probe_site_order_${i + 1}`, siteProbe(o), i ? ALWAYS : undefined]),
    ['Site_order', compose(`@${siteOrderExpression()}`), ALWAYS],
    [
      'Search_refused',
      {
        type: 'If',
        expression: { equals: ["@outputs('Site_order')", 'refused'] },
        actions: { Mark_search_refused: { ...http('POST', "outputs('Home')", settingsItem, { headers: merge, body: `{"ARLastRun":"Failed @{${now}} UTC: SharePoint search refused every way of listing the sites. @{${probeErrors()}}"}` }), runAfter: {} } },
        else: { actions: {} }
      }
    ],
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
        body: `{"ARLastRun":"Finished @{${now}} UTC: @{length(variables('Sites'))} site collections, @{length(body('Subjects'))} people and groups, @{length(body('Rows_with_access'))} rows, @{length(body('Unread_rows'))} not read. File: access-report-@{outputs('File_stamp')}.csv"}`
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
