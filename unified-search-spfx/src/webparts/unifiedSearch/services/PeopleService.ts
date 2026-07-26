import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { IPersonResult, IPersonOrg, PresenceAvailability } from '../models';

interface IGraphUser {
  id?: string;
  displayName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  mobilePhone?: string;
}

const USER_SELECT =
  'id,displayName,jobTitle,department,officeLocation,mail,userPrincipalName,businessPhones,mobilePhone';

export class PeopleService {

  private _clientPromise: Promise<MSGraphClientV3>;
  private _photoCache: { [userId: string]: string | undefined } = {};

  constructor(factory: MSGraphClientFactory) {
    this._clientPromise = factory.getClient('3');
  }

  /**
   * Directory search over users. Delegated (User.Read.All), so the caller only
   * ever sees directory entries they are allowed to read.
   */
  public async search(query: string, top: number): Promise<IPersonResult[]> {
    const client = await this._clientPromise;
    const escaped = query.replace(/"/g, '\\"');
    const search =
      `"displayName:${escaped}" OR "jobTitle:${escaped}" OR "department:${escaped}" OR "mail:${escaped}"`;
    const response: { value?: IGraphUser[] } = await client
      .api('/users')
      .header('ConsistencyLevel', 'eventual')
      .search(search)
      .filter('accountEnabled eq true')
      .select(USER_SELECT)
      .top(top)
      .count(true)
      .get();
    return (response.value || []).map(user => this._toPerson(user));
  }

  public async getPresence(people: IPersonResult[]): Promise<{ [userId: string]: PresenceAvailability }> {
    const ids = people.map(person => person.id).filter(id => !!id);
    if (ids.length === 0) {
      return {};
    }
    const client = await this._clientPromise;
    const response: { value?: { id?: string; availability?: string }[] } = await client
      .api('/communications/getPresencesByUserId')
      .post({ ids });
    const map: { [userId: string]: PresenceAvailability } = {};
    (response.value || []).forEach(presence => {
      if (presence.id) {
        map[presence.id] = (presence.availability || 'PresenceUnknown') as PresenceAvailability;
      }
    });
    return map;
  }

  public async getPhoto(userId: string): Promise<string | undefined> {
    if (Object.prototype.hasOwnProperty.call(this._photoCache, userId)) {
      return this._photoCache[userId];
    }
    try {
      const client = await this._clientPromise;
      const blob: Blob = await client
        .api(`/users/${userId}/photos/48x48/$value`)
        .responseType('blob' as never)
        .get();
      const url = URL.createObjectURL(blob);
      this._photoCache[userId] = url;
      return url;
    } catch {
      this._photoCache[userId] = undefined;
      return undefined;
    }
  }

  /** Manager + direct reports for the person card's mini org chart. */
  public async getOrg(userId: string): Promise<IPersonOrg> {
    const client = await this._clientPromise;
    const [manager, reports] = await Promise.all([
      client.api(`/users/${userId}/manager`).select(USER_SELECT).get()
        .catch(() => undefined),
      client.api(`/users/${userId}/directReports`).select(USER_SELECT).get()
        .catch((): { value: IGraphUser[] } => ({ value: [] }))
    ]);
    return {
      manager: manager ? this._toPerson(manager as IGraphUser) : undefined,
      directReports: ((reports as { value?: IGraphUser[] }).value || []).map(user => this._toPerson(user))
    };
  }

  private _toPerson(user: IGraphUser): IPersonResult {
    return {
      id: user.id || '',
      displayName: user.displayName || '',
      jobTitle: user.jobTitle,
      department: user.department,
      officeLocation: user.officeLocation,
      mail: user.mail || user.userPrincipalName,
      userPrincipalName: user.userPrincipalName,
      businessPhone: user.businessPhones && user.businessPhones.length > 0 ? user.businessPhones[0] : undefined,
      mobilePhone: user.mobilePhone
    };
  }
}
