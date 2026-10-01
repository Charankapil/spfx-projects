import { SpClient } from './SpClient';
import { IPerson } from '../models/types';

interface IPickerEntity {
  Key: string;
  DisplayText: string;
  EntityData?: { Email?: string };
}

/** People search and resolution through SharePoint's own people picker API. No Graph. */
export class PeopleService {
  private readonly sp: SpClient;
  private readonly ensured: { [key: string]: IPerson } = {};
  private readonly byId: Map<number, IPerson> = new Map();
  private me: IPerson | null = null;

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  public async currentUser(): Promise<IPerson> {
    if (!this.me) {
      const u = await this.sp.get<{ Id: number; Title: string; Email: string }>('web/currentuser?$select=Id,Title,Email');
      this.me = { id: u.Id, title: u.Title, email: u.Email };
    }
    return this.me;
  }

  /** Search people in the directory. Returns picker keys; call ensure() to get site user ids. */
  public async search(query: string, max: number = 8): Promise<{ key: string; title: string; email: string }[]> {
    if (!query || query.trim().length < 2) {
      return [];
    }
    const res = await this.sp.post<{ value?: string; d?: { ClientPeoplePickerSearchUser: string } }>(
      'SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser',
      {
        queryParams: {
          __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
          AllowEmailAddresses: false,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: max,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: query.trim()
        }
      }
    );
    const raw = (res && (res.value || (res.d && res.d.ClientPeoplePickerSearchUser))) || '[]';
    const entities: IPickerEntity[] = JSON.parse(raw);
    return entities.map(e => ({
      key: e.Key,
      title: e.DisplayText,
      email: (e.EntityData && e.EntityData.Email) || ''
    }));
  }

  /**
   * Names and emails for site user ids, from the site's user list. Cached; asks for up to 20 at a time.
   * Ids that can't be found (for example a deleted user) come back as "User 12".
   */
  public async resolve(ids: number[]): Promise<Map<number, IPerson>> {
    const missing = ids.filter(id => !this.byId.has(id));
    for (let i = 0; i < missing.length; i += 20) {
      const chunk = missing.slice(i, i + 20);
      const filter = chunk.map(id => `Id eq ${id}`).join(' or ');
      try {
        const res = await this.sp.get<{ value: { Id: number; Title: string; Email: string }[] }>(
          `web/siteusers?$select=Id,Title,Email&$filter=${encodeURIComponent(filter)}`
        );
        res.value.forEach(u => this.byId.set(u.Id, { id: u.Id, title: u.Title, email: u.Email || '' }));
      } catch {
        // Fall back to "User 12" labels rather than failing the whole board.
      }
      chunk.forEach(id => {
        if (!this.byId.has(id)) {
          this.byId.set(id, { id, title: 'User ' + id, email: '' });
        }
      });
    }
    const out = new Map<number, IPerson>();
    ids.forEach(id => out.set(id, this.byId.get(id) as IPerson));
    return out;
  }

  /**
   * The signed-in user's reports, from their SharePoint user profile (org data synced from
   * Entra ID). Uses SharePoint's own profile API: no Graph, no app registration.
   * `everyone` includes reports of reports. Returns account names (login names).
   */
  public async myReportLogins(everyone: boolean): Promise<string[]> {
    type Logins = string[] | { results: string[] } | null | undefined;
    const res = await this.sp.get<{ AccountName?: string; DirectReports?: Logins; ExtendedReports?: Logins }>(
      'SP.UserProfiles.PeopleManager/GetMyProperties?$select=AccountName,DirectReports,ExtendedReports'
    );
    const list = (raw: Logins): string[] => (Array.isArray(raw) ? raw : raw && Array.isArray(raw.results) ? raw.results : []);
    const me = String(res.AccountName || '').toLowerCase();
    const logins = list(everyone ? res.ExtendedReports : res.DirectReports);
    return logins.filter((l, i) => l && l.toLowerCase() !== me && logins.indexOf(l) === i);
  }

  /**
   * Site users for login names. People who were never added to this site are not returned:
   * they cannot be in any board's People column here.
   */
  public async siteUsersByLogin(logins: string[]): Promise<IPerson[]> {
    const out: IPerson[] = [];
    for (let i = 0; i < logins.length; i += 15) {
      const chunk = logins.slice(i, i + 15);
      const filter = chunk.map(l => `LoginName eq '${l.replace(/'/g, "''")}'`).join(' or ');
      const res = await this.sp.get<{ value: { Id: number; Title: string; Email: string }[] }>(
        `web/siteusers?$select=Id,Title,Email&$filter=${encodeURIComponent(filter)}`
      );
      res.value.forEach(u => {
        const person = { id: u.Id, title: u.Title, email: u.Email || '' };
        this.byId.set(u.Id, person);
        out.push(person);
      });
    }
    return out;
  }

  /** Make sure the person exists in the site's user list and return their site user id. */
  public async ensure(key: string): Promise<IPerson> {
    if (this.ensured[key]) {
      return this.ensured[key];
    }
    const u = await this.sp.post<{ Id: number; Title: string; Email: string }>('web/ensureuser', { logonName: key });
    const person: IPerson = { id: u.Id, title: u.Title, email: u.Email };
    this.ensured[key] = person;
    this.byId.set(person.id, person);
    return person;
  }
}

export function initials(name: string): string {
  const parts = name.replace(/\(.*?\)/g, '').trim().split(/\s+/).filter(p => p.length > 0);
  if (parts.length === 0) {
    return '?';
  }
  if (parts.length === 1) {
    return parts[0].substring(0, 2).toUpperCase();
  }
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

/** Stable avatar colour per person. */
export function personColor(id: number): string {
  const colors = ['#0073ea', '#00854d', '#a25ddc', '#e2445c', '#fdab3d', '#037f4c', '#784bd1', '#bb3354', '#0086c0', '#7f5347'];
  return colors[Math.abs(id) % colors.length];
}
