import { runMigrations, getSetupStatus, MIGRATIONS, SCHEMA_VERSION } from '../Provisioner';
import { SpClient, SpError } from '../SpClient';

/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * In-memory fake of the few SharePoint REST endpoints the provisioner uses.
 * Lets us check that setup creates everything once and is safe to re-run.
 */
class FakeSp {
  public webUrl: string = 'https://contoso.sharepoint.com/sites/team';
  public lists: { [title: string]: { Id: string; fields: string[]; items: any[]; settings: any } } = {};
  public calls: string[] = [];
  public perms: { High: string; Low: string } = { High: '2147483647', Low: '4294967295' };

  private byId(id: string): { Id: string; fields: string[]; items: any[]; settings: any } {
    const t = Object.keys(this.lists).filter(k => this.lists[k].Id === id)[0];
    if (!t) {
      throw new SpError('List not found', 404, '');
    }
    return this.lists[t];
  }

  public async get(path: string): Promise<any> {
    this.calls.push('GET ' + path);
    if (path === 'web/EffectiveBasePermissions') {
      return this.perms;
    }
    let m = /^web\/lists\/getbytitle\('([^']+)'\)/.exec(path);
    if (m) {
      const list = this.lists[decodeURIComponent(m[1])];
      if (!list) {
        throw new SpError('List does not exist', 404, '');
      }
      return { Id: list.Id, Title: m[1] };
    }
    m = /^web\/lists\(guid'([^']+)'\)\/fields\/getbyinternalnameortitle\('([^']+)'\)/.exec(path);
    if (m) {
      const list = this.byId(m[1]);
      if (list.fields.indexOf(decodeURIComponent(m[2])) < 0) {
        throw new SpError('Field not found', 404, '');
      }
      return { InternalName: m[2] };
    }
    m = /^web\/lists\(guid'([^']+)'\)\/items\?.*\$filter=Title eq 'schema'/.exec(path);
    if (m) {
      const list = this.byId(m[1]);
      if (list.fields.indexOf('WB_Version') < 0) {
        throw new SpError('Column WB_Version does not exist', 400, '');
      }
      return { value: list.items.filter(i => i.Title === 'schema') };
    }
    m = /^web\/lists\(guid'([^']+)'\)\?\$select=ListItemEntityTypeFullName/.exec(path);
    if (m) {
      return { ListItemEntityTypeFullName: 'SP.Data.TestListItem' };
    }
    throw new Error('Unexpected GET ' + path);
  }

  public async post(path: string, body?: any): Promise<any> {
    this.calls.push('POST ' + path);
    if (path === 'web/lists') {
      const id = 'id-' + Object.keys(this.lists).length;
      this.lists[body.Title] = { Id: id, fields: ['Title'], items: [], settings: {} };
      return { Id: id, Title: body.Title };
    }
    let m = /^web\/lists\(guid'([^']+)'\)\/fields\/createfieldasxml$/.exec(path);
    if (m) {
      // Real SharePoint: with the internal name hint (Options 8) the internal name comes from DisplayName, not Name.
      const xml: string = body.parameters.SchemaXml;
      const source = body.parameters.Options === 8 ? /DisplayName="([^"]+)"/ : /Name="([^"]+)"/;
      this.byId(m[1]).fields.push((source.exec(xml) as RegExpExecArray)[1]);
      return {};
    }
    m = /^web\/lists\(guid'([^']+)'\)\/items$/.exec(path);
    if (m) {
      const list = this.byId(m[1]);
      const row = { ...body, Id: list.items.length + 1 };
      list.items.push(row);
      return row;
    }
    throw new Error('Unexpected POST ' + path);
  }

  public async merge(path: string, body: any): Promise<string> {
    this.calls.push('MERGE ' + path);
    let m = /^web\/lists\(guid'([^']+)'\)$/.exec(path);
    if (m) {
      Object.assign(this.byId(m[1]).settings, body);
      return '';
    }
    if (/^web\/lists\(guid'([^']+)'\)\/fields\//.test(path)) {
      return '';
    }
    m = /^web\/lists\(guid'([^']+)'\)\/items\((\d+)\)$/.exec(path);
    if (m) {
      const row = this.byId(m[1]).items.filter(i => i.Id === Number(m![2]))[0];
      Object.assign(row, body);
      return '';
    }
    throw new Error('Unexpected MERGE ' + path);
  }
}

describe('provisioner', () => {
  it('reports a fresh site as not set up, and whether the user can set it up', async () => {
    const fake = new FakeSp();
    const status = await getSetupStatus(fake as unknown as SpClient);
    expect(status).toEqual({ installedVersion: 0, requiredVersion: SCHEMA_VERSION, canSetup: true });
    fake.perms = { High: '0', Low: '1' };
    expect((await getSetupStatus(fake as unknown as SpClient)).canSetup).toBe(false);
  });

  it('creates the site lists once and records the schema version', async () => {
    const fake = new FakeSp();
    const sp = fake as unknown as SpClient;
    const progress: string[] = [];
    await runMigrations(sp, 0, (i, s) => progress.push(`${i}:${s}`));

    expect(Object.keys(fake.lists).sort()).toEqual(['WB_Boards', 'WB_Meta', 'WB_UserPrefs']);
    expect(fake.lists.WB_Boards.fields).toEqual(expect.arrayContaining(['WB_Key', 'WB_Config', 'WB_BoardOwners', 'WB_Privacy']));
    // Fields exist under their WB_ internal names (the bug was creating "Version" instead of "WB_Version").
    expect(fake.lists.WB_Meta.fields).toEqual(expect.arrayContaining(['WB_Version', 'WB_Settings']));
    expect(fake.lists.WB_Meta.fields).not.toContain('Version');
    expect(fake.lists.WB_UserPrefs.fields).toContain('WB_Prefs');
    expect(fake.lists.WB_UserPrefs.settings).toEqual(expect.objectContaining({ ReadSecurity: 2, WriteSecurity: 2, Hidden: true }));
    expect(fake.lists.WB_Meta.items).toEqual([expect.objectContaining({ Title: 'schema', WB_Version: SCHEMA_VERSION })]);
    expect(progress[progress.length - 1]).toBe(`${MIGRATIONS.length - 1}:done`);
    expect((await getSetupStatus(sp)).installedVersion).toBe(SCHEMA_VERSION);

    // Running again creates nothing new.
    const creates = (): number => fake.calls.filter(c => c === 'POST web/lists' || c.indexOf('createfieldasxml') >= 0).length;
    const before = creates();
    await runMigrations(sp, 0, () => undefined);
    expect(creates()).toBe(before);
    expect(fake.lists.WB_Meta.items).toHaveLength(1);
  });

  it('skips migrations that are already installed', async () => {
    const fake = new FakeSp();
    const states: string[] = [];
    await runMigrations(fake as unknown as SpClient, SCHEMA_VERSION, (i, s) => states.push(s));
    expect(states.every(s => s === 'done')).toBe(true);
    expect(Object.keys(fake.lists)).toHaveLength(0);
  });
});
