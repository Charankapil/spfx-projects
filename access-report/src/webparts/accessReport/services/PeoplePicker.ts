import { entityOf, SPClient, trimSlash } from '../core/SPClient';
import { isM365Login, ISubject, normaliseLogin, subjectKind } from './permissions';

/**
 * Finds people and groups with SharePoint's own people picker service - the one behind every
 * "Share" dialog. It answers as the signed-in user and returns the SharePoint login name
 * (claim) the report needs, so no Microsoft Graph permission is involved.
 */

interface IPickerEntity {
  Key?: string;
  DisplayText?: string;
  EntityType?: string;
  IsResolved?: boolean;
  EntityData?: { Email?: string; PrincipalType?: string };
}

export interface IFound extends ISubject {
  kind: string;
}

/**
 * Picker entities -> subjects: unresolved entries and SharePoint groups are left out, and a
 * Microsoft 365 group's "Owners" entry folds into the group (its owners are checked with it).
 */
export function fromPicker(entities: IPickerEntity[]): IFound[] {
  const out: IFound[] = [];
  const add = (e: IPickerEntity, ownersEntry: boolean): void => {
    const key = String(e.Key || '');
    const login = normaliseLogin(key);
    if (out.some((o) => o.login.toLowerCase() === login.toLowerCase())) {
      return;
    }
    const name = String(e.DisplayText || key).replace(ownersEntry ? /\s+owners$/i : /$^/, '');
    const mail = (e.EntityData && e.EntityData.Email) || '';
    out.push({ login, label: mail && name.toLowerCase().indexOf(mail.toLowerCase()) < 0 ? `${name} <${mail}>` : name, kind: subjectKind(login) });
  };
  // A claim always has "|"; SharePoint groups come back as a bare name, and are per site anyway.
  const usable = entities.filter((e) => String(e.Key || '').indexOf('|') >= 0 && e.EntityType !== 'SPGroup');
  const isOwners = (e: IPickerEntity): boolean => isM365Login(String(e.Key)) && /_o$/i.test(String(e.Key));
  usable.filter((e) => !isOwners(e)).forEach((e) => add(e, false));
  usable.filter(isOwners).forEach((e) => add(e, true));
  return out;
}

export class PeoplePicker {
  constructor(private client: SPClient, private homeWebUrl: string) {}

  public async search(text: string): Promise<IFound[]> {
    const q = text.trim();
    if (!q) {
      return [];
    }
    const json = await this.client.post<unknown>(
      `${trimSlash(this.homeWebUrl)}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
      {
        queryParams: {
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: 25,
          // 15 = every source; 5 = users + security groups (Microsoft 365 groups are returned as those).
          PrincipalSource: 15,
          PrincipalType: 5,
          QueryString: q
        }
      },
      { keepCache: true }
    );
    const o = entityOf<{ value?: unknown; ClientPeoplePickerSearchUser?: unknown }>(json) || {};
    const raw = o.value !== undefined ? o.value : o.ClientPeoplePickerSearchUser;
    const entities = typeof raw === 'string' ? (JSON.parse(raw) as IPickerEntity[]) : Array.isArray(raw) ? (raw as IPickerEntity[]) : [];
    return fromPicker(entities);
  }
}
