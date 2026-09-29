import { SpClient, SpError } from './SpClient';
import { ensureField, ensureList, findList, itemEntityType } from './lists';
import { IBasePermissions, hasPermission, PermissionKind } from './permissions';

export const LIST_META = 'WB_Meta';
export const LIST_BOARDS = 'WB_Boards';
export const LIST_PREFS = 'WB_UserPrefs';

export interface IMigration {
  version: number;
  title: string;
  run: (sp: SpClient) => Promise<void>;
}

/**
 * Numbered, idempotent migrations. Each one checks before it creates, so an interrupted
 * setup or upgrade can run again. Add new migrations to the end; never change old ones.
 */
export const MIGRATIONS: IMigration[] = [
  {
    version: 1,
    title: 'Settings list (WB_Meta)',
    run: async sp => {
      const list = await ensureList(sp, LIST_META, { description: 'Work Boards settings and schema version.', hidden: true });
      await ensureField(sp, list.Id, 'WB_Version', '<Field Type="Number" Name="WB_Version" StaticName="WB_Version" DisplayName="Version" />');
      await ensureField(sp, list.Id, 'WB_Settings', '<Field Type="Note" Name="WB_Settings" StaticName="WB_Settings" DisplayName="Settings" NumLines="6" RichText="FALSE" />');
    }
  },
  {
    version: 2,
    title: 'Board registry (WB_Boards)',
    run: async sp => {
      const list = await ensureList(sp, LIST_BOARDS, { description: 'Work Boards: one row per board.', hidden: true, versioning: true });
      const fields: [string, string][] = [
        ['WB_Key', '<Field Type="Text" Name="WB_Key" StaticName="WB_Key" DisplayName="Board key" MaxLength="16" Indexed="TRUE" EnforceUniqueValues="TRUE" />'],
        ['WB_Description', '<Field Type="Note" Name="WB_Description" StaticName="WB_Description" DisplayName="Description" NumLines="3" RichText="FALSE" />'],
        ['WB_Folder', '<Field Type="Text" Name="WB_Folder" StaticName="WB_Folder" DisplayName="Folder" MaxLength="100" />'],
        ['WB_Color', '<Field Type="Text" Name="WB_Color" StaticName="WB_Color" DisplayName="Colour" MaxLength="16" />'],
        ['WB_Privacy', '<Field Type="Choice" Name="WB_Privacy" StaticName="WB_Privacy" DisplayName="Privacy" Format="Dropdown"><CHOICES><CHOICE>Main</CHOICE><CHOICE>Private</CHOICE></CHOICES><Default>Main</Default></Field>'],
        ['WB_ItemsListId', '<Field Type="Text" Name="WB_ItemsListId" StaticName="WB_ItemsListId" DisplayName="Items list id" MaxLength="40" />'],
        ['WB_UpdatesListId', '<Field Type="Text" Name="WB_UpdatesListId" StaticName="WB_UpdatesListId" DisplayName="Updates list id" MaxLength="40" />'],
        ['WB_BoardOwners', '<Field Type="UserMulti" Name="WB_BoardOwners" StaticName="WB_BoardOwners" DisplayName="Board owners" Mult="TRUE" UserSelectionMode="PeopleOnly" />'],
        ['WB_Config', '<Field Type="Note" Name="WB_Config" StaticName="WB_Config" DisplayName="Configuration" NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />'],
        ['WB_Archived', '<Field Type="Boolean" Name="WB_Archived" StaticName="WB_Archived" DisplayName="Archived"><Default>0</Default></Field>']
      ];
      for (const [name, xml] of fields) {
        await ensureField(sp, list.Id, name, xml);
      }
    }
  },
  {
    version: 3,
    title: 'Personal settings (WB_UserPrefs, own rows only)',
    run: async sp => {
      const list = await ensureList(sp, LIST_PREFS, {
        description: 'Work Boards: favourites and recent boards. Each person sees only their own row.',
        hidden: true,
        ownItemsOnly: true
      });
      await ensureField(sp, list.Id, 'WB_Prefs', '<Field Type="Note" Name="WB_Prefs" StaticName="WB_Prefs" DisplayName="Preferences" NumLines="6" RichText="FALSE" />');
    }
  }
];

export const SCHEMA_VERSION = MIGRATIONS[MIGRATIONS.length - 1].version;

export interface ISetupStatus {
  installedVersion: number;
  requiredVersion: number;
  canSetup: boolean;
}

interface IMetaRow {
  Id: number;
  WB_Version: number | null;
}

async function readMetaRow(sp: SpClient): Promise<IMetaRow | null> {
  const list = await findList(sp, LIST_META);
  if (!list) {
    return null;
  }
  try {
    const rows = await sp.get<{ value: IMetaRow[] }>(`web/lists(guid'${list.Id}')/items?$select=Id,WB_Version&$filter=Title eq 'schema'&$top=1`);
    return rows.value[0] || null;
  } catch (e) {
    if (e instanceof SpError && (e.status === 400 || e.isNotFound)) {
      // WB_Version field not created yet (setup interrupted).
      return null;
    }
    throw e;
  }
}

export async function getSetupStatus(sp: SpClient): Promise<ISetupStatus> {
  const [row, perms] = await Promise.all([
    readMetaRow(sp),
    sp.get<IBasePermissions>('web/EffectiveBasePermissions')
  ]);
  return {
    installedVersion: row && row.WB_Version ? row.WB_Version : 0,
    requiredVersion: SCHEMA_VERSION,
    canSetup: hasPermission(perms, PermissionKind.ManageWeb)
  };
}

async function writeVersion(sp: SpClient, version: number): Promise<void> {
  const list = await findList(sp, LIST_META);
  if (!list) {
    throw new Error('The WB_Meta list is missing.');
  }
  const row = await readMetaRow(sp);
  const type = await itemEntityType(sp, list.Id);
  if (row) {
    await sp.merge(`web/lists(guid'${list.Id}')/items(${row.Id})`, { __metadata: { type }, WB_Version: version });
  } else {
    await sp.post(`web/lists(guid'${list.Id}')/items`, { __metadata: { type }, Title: 'schema', WB_Version: version });
  }
}

export type MigrationProgress = (index: number, state: 'running' | 'done' | 'error', message?: string) => void;

/** Run every migration above `fromVersion`, recording progress after each one. */
export async function runMigrations(sp: SpClient, fromVersion: number, onProgress: MigrationProgress): Promise<void> {
  for (let i = 0; i < MIGRATIONS.length; i++) {
    const m = MIGRATIONS[i];
    if (m.version <= fromVersion) {
      onProgress(i, 'done');
      continue;
    }
    onProgress(i, 'running');
    try {
      await m.run(sp);
      await writeVersion(sp, m.version);
      onProgress(i, 'done');
    } catch (e) {
      onProgress(i, 'error', (e as Error).message);
      throw e;
    }
  }
}
