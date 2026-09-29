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

  /** Make sure the person exists in the site's user list and return their site user id. */
  public async ensure(key: string): Promise<IPerson> {
    if (this.ensured[key]) {
      return this.ensured[key];
    }
    const u = await this.sp.post<{ Id: number; Title: string; Email: string }>('web/ensureuser', { logonName: key });
    const person: IPerson = { id: u.Id, title: u.Title, email: u.Email };
    this.ensured[key] = person;
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
