import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { CacheService } from './CacheService';

export interface IGraphUser {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
  companyName?: string;
  employeeId?: string;
  city?: string;
  country?: string;
  state?: string;
  postalCode?: string;
  streetAddress?: string;
  preferredLanguage?: string;
  aboutMe?: string;
  skills?: string[];
  interests?: string[];
  schools?: string[];
  pastProjects?: string[];
  responsibilities?: string[];
  mySite?: string;
  preferredName?: string;
  birthday?: string;
  hireDate?: string;
  accountEnabled?: boolean;
  manager?: IGraphUser;
  directReports?: IGraphUser[];
}

export interface IGraphGroup {
  id: string;
  displayName: string;
  mail?: string;
  mailNickname?: string;
}

export interface IGraphPresence {
  id: string;
  availability: string;
  activity: string;
}

export interface IGraphInsight {
  id: string;
  lastUsed?: { lastAccessedDateTime?: string; lastModifiedDateTime?: string };
  resourceVisualization: {
    title: string;
    type: string;
    mediaType?: string;
    previewImageUrl?: string;
    containerDisplayName?: string;
    containerWebUrl?: string;
    containerType?: string;
  };
  resourceReference: {
    webUrl: string;
    id?: string;
    type?: string;
  };
}

export interface IGraphSite {
  id: string;
  displayName: string;
  name?: string;
  webUrl: string;
  description?: string;
  siteCollection?: { hostname: string };
}

export interface IGraphDriveItem {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  folder?: { childCount: number };
  file?: { mimeType?: string };
  package?: { type: string };
  parentReference?: { driveId?: string; id?: string; path?: string };
  remoteItem?: { webUrl?: string; name?: string; lastModifiedDateTime?: string; shared?: { sharedBy?: { user?: { displayName?: string } } } };
  shared?: { sharedBy?: { user?: { displayName?: string } } };
}

export interface IGraphDrive {
  id: string;
  webUrl: string;
  quota?: { total?: number; used?: number; remaining?: number; state?: string };
}

export interface IGraphEvent {
  id: string;
  subject: string;
  webLink: string;
  isAllDay: boolean;
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string };
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  organizer?: { emailAddress?: { name?: string } };
  responseStatus?: { response?: string };
  showAs?: string;
}

export interface IGraphTodoTask {
  id: string;
  title: string;
  status: string;
  importance: string;
  dueDateTime?: { dateTime: string; timeZone: string };
  listId?: string;
  listName?: string;
}

export interface IGraphTodoList {
  id: string;
  displayName: string;
  wellknownListName?: string;
}

export interface IGraphPlannerTask {
  id: string;
  title: string;
  percentComplete: number;
  dueDateTime?: string;
  planId?: string;
  priority?: number;
  planTitle?: string;
}

export interface IGraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  isRead: boolean;
  webLink: string;
  importance?: string;
  hasAttachments?: boolean;
  from?: { emailAddress?: { name?: string; address?: string } };
}

export interface IGraphTeam {
  id: string;
  displayName: string;
  description?: string;
  webUrl?: string;
  isArchived?: boolean;
}

export interface IGraphPerson {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  userPrincipalName?: string;
  scoredEmailAddresses?: { address: string; relevanceScore: number }[];
  personType?: { class?: string; subclass?: string };
}

export interface IGraphAppRoleAssignment {
  id: string;
  resourceDisplayName: string;
  resourceId: string;
  principalType?: string;
}

export interface IGraphServicePrincipal {
  id: string;
  appId: string;
  displayName: string;
  homepage?: string;
  loginUrl?: string;
  info?: { logoUrl?: string; marketingUrl?: string };
  tags?: string[];
  accountEnabled?: boolean;
}

export interface IGraphCollection<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export const USER_SELECT: string =
  'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,companyName,employeeId,city,country,state,postalCode,streetAddress,preferredLanguage,accountEnabled';

export const USER_PROFILE_SELECT: string =
  `${USER_SELECT},aboutMe,skills,interests,schools,pastProjects,responsibilities,mySite,preferredName,birthday,hireDate`;

/**
 * Thin, typed wrapper around MSGraphClientV3 with caching and error normalisation.
 */
export class GraphService {
  private _client: MSGraphClientV3 | undefined;
  private readonly _context: WebPartContext;
  public readonly cache: CacheService;
  private readonly _photoCache: Map<string, Promise<string | undefined>> = new Map();

  constructor(context: WebPartContext, cache: CacheService) {
    this._context = context;
    this.cache = cache;
  }

  public get tenantId(): string {
    return this._context.pageContext.aadInfo?.tenantId?.toString() || '';
  }

  private async _getClient(): Promise<MSGraphClientV3> {
    if (!this._client) {
      this._client = await this._context.msGraphClientFactory.getClient('3');
    }
    return this._client;
  }

  /** GET with optional caching. */
  public async get<T>(path: string, options?: { cacheKey?: string; headers?: Record<string, string>; noCache?: boolean }): Promise<T> {
    const cacheKey: string = options?.cacheKey || `get:${path}`;
    if (!options?.noCache) {
      const cached: T | undefined = this.cache.get<T>(cacheKey);
      if (cached !== undefined) {
        return cached;
      }
    }
    const client: MSGraphClientV3 = await this._getClient();
    let request = client.api(path);
    if (options?.headers) {
      request = request.headers(options.headers);
    }
    const result: T = await request.get();
    if (!options?.noCache) {
      this.cache.put(cacheKey, result);
    }
    return result;
  }

  /** GET and follow @odata.nextLink pages until `max` items are collected. */
  public async getAll<T>(path: string, max: number, options?: { headers?: Record<string, string>; noCache?: boolean }): Promise<T[]> {
    const cacheKey: string = `all:${path}:${max}`;
    if (!options?.noCache) {
      const cached: T[] | undefined = this.cache.get<T[]>(cacheKey);
      if (cached) {
        return cached;
      }
    }
    const client: MSGraphClientV3 = await this._getClient();
    const items: T[] = [];
    let next: string | undefined = path;
    while (next && items.length < max) {
      let request = client.api(next);
      if (options?.headers) {
        request = request.headers(options.headers);
      }
      const page: IGraphCollection<T> = await request.get();
      items.push(...(page.value || []));
      next = page['@odata.nextLink'];
    }
    const sliced: T[] = items.slice(0, max);
    if (!options?.noCache) {
      this.cache.put(cacheKey, sliced);
    }
    return sliced;
  }

  public async patch<T>(path: string, body: unknown): Promise<T> {
    const client: MSGraphClientV3 = await this._getClient();
    return client.api(path).patch(body);
  }

  public async post<T>(path: string, body: unknown): Promise<T> {
    const client: MSGraphClientV3 = await this._getClient();
    return client.api(path).post(body);
  }

  public async put<T>(path: string, body: unknown, contentType: string = 'application/json'): Promise<T> {
    const client: MSGraphClientV3 = await this._getClient();
    return client.api(path).header('Content-Type', contentType).put(body);
  }

  public async delete(path: string): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    await client.api(path).delete();
  }

  // ---------------------------------------------------------------------------
  // Users & people
  // ---------------------------------------------------------------------------

  public getMe(): Promise<IGraphUser> {
    return this.get<IGraphUser>(`/me?$select=${USER_PROFILE_SELECT}`);
  }

  public async updateMe(patch: Partial<IGraphUser>): Promise<void> {
    await this.patch('/me', patch);
    this.cache.remove(`get:/me?$select=${USER_PROFILE_SELECT}`);
  }

  public getUser(id: string): Promise<IGraphUser> {
    return this.get<IGraphUser>(`/users/${encodeURIComponent(id)}?$select=${USER_SELECT}`);
  }

  public async getMyManager(): Promise<IGraphUser | undefined> {
    try {
      return await this.get<IGraphUser>(`/me/manager?$select=${USER_SELECT}`);
    } catch (e) {
      if (isNotFound(e)) {
        return undefined;
      }
      throw e;
    }
  }

  public getDirectReports(userId: string = 'me'): Promise<IGraphUser[]> {
    const base: string = userId === 'me' ? '/me' : `/users/${encodeURIComponent(userId)}`;
    return this.getAll<IGraphUser>(`${base}/directReports?$select=${USER_SELECT}&$top=50`, 200);
  }

  public async getManagerChain(): Promise<IGraphUser[]> {
    // $levels=max requires the eventual consistency header.
    const me: IGraphUser = await this.get<IGraphUser>(
      `/me?$select=id,displayName&$expand=manager($levels=max;$select=${USER_SELECT})`,
      { headers: { ConsistencyLevel: 'eventual' }, cacheKey: 'managerChain' }
    );
    const chain: IGraphUser[] = [];
    let cursor: IGraphUser | undefined = me.manager;
    let guard: number = 0;
    while (cursor && guard < 25) {
      chain.push(cursor);
      cursor = cursor.manager;
      guard++;
    }
    return chain;
  }

  public getPeople(top: number): Promise<IGraphPerson[]> {
    return this.getAll<IGraphPerson>(`/me/people?$top=${top}&$filter=personType/class eq 'Person'`, top);
  }

  /** Groups the user belongs to, including nested membership. */
  public getMyGroups(): Promise<IGraphGroup[]> {
    return this.getAll<IGraphGroup>('/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName,mail,mailNickname&$top=999', 2000, {
      headers: { ConsistencyLevel: 'eventual' }
    });
  }

  public async hasDirectReports(): Promise<boolean> {
    const page: IGraphCollection<{ id: string }> = await this.get('/me/directReports?$select=id&$top=1', { cacheKey: 'me:hasReports' });
    return (page.value || []).length > 0;
  }

  public async searchUsers(query: string, top: number = 10): Promise<IGraphUser[]> {
    const q: string = query.replace(/'/g, "''");
    return this.getAll<IGraphUser>(
      `/users?$search="displayName:${q}" OR "mail:${q}"&$select=${USER_SELECT}&$top=${top}&$orderby=displayName`,
      top,
      { headers: { ConsistencyLevel: 'eventual' } }
    );
  }

  public async getPresence(userIds: string[]): Promise<IGraphPresence[]> {
    if (userIds.length === 0) {
      return [];
    }
    const result: IGraphCollection<IGraphPresence> = await this.post('/communications/getPresencesByUserId', { ids: userIds.slice(0, 650) });
    return result.value || [];
  }

  /**
   * Returns an object URL for the user's photo or undefined when no photo exists.
   */
  public getPhoto(userId: string, size: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' = '96x96'): Promise<string | undefined> {
    const key: string = `${userId}:${size}`;
    const existing: Promise<string | undefined> | undefined = this._photoCache.get(key);
    if (existing) {
      return existing;
    }
    const promise: Promise<string | undefined> = (async (): Promise<string | undefined> => {
      try {
        const client: MSGraphClientV3 = await this._getClient();
        const base: string = userId === 'me' ? '/me' : `/users/${encodeURIComponent(userId)}`;
        const blob: Blob = await client.api(`${base}/photos/${size}/$value`).responseType('blob' as never).get();
        if (!blob || blob.size === 0) {
          return undefined;
        }
        return URL.createObjectURL(blob);
      } catch {
        return undefined;
      }
    })();
    this._photoCache.set(key, promise);
    return promise;
  }

  // ---------------------------------------------------------------------------
  // Files, sites & insights
  // ---------------------------------------------------------------------------

  public getUsedInsights(top: number): Promise<IGraphInsight[]> {
    return this.getAll<IGraphInsight>(`/me/insights/used?$top=${top}`, top);
  }

  public getTrendingInsights(top: number): Promise<IGraphInsight[]> {
    return this.getAll<IGraphInsight>(`/me/insights/trending?$top=${top}`, top);
  }

  public getSharedInsights(top: number): Promise<IGraphInsight[]> {
    return this.getAll<IGraphInsight>(`/me/insights/shared?$top=${top}`, top);
  }

  public getFollowedSites(top: number): Promise<IGraphSite[]> {
    return this.getAll<IGraphSite>(`/me/followedSites?$top=${top}`, top);
  }

  public searchSites(query: string, top: number): Promise<IGraphSite[]> {
    return this.getAll<IGraphSite>(`/sites?search=${encodeURIComponent(query)}&$top=${top}`, top);
  }

  public getMyDrive(): Promise<IGraphDrive> {
    return this.get<IGraphDrive>('/me/drive?$select=id,webUrl,quota');
  }

  public getDriveChildren(itemId: string | undefined, top: number): Promise<IGraphDriveItem[]> {
    const base: string = itemId ? `/me/drive/items/${itemId}/children` : '/me/drive/root/children';
    return this.getAll<IGraphDriveItem>(
      `${base}?$top=${top}&$orderby=name&$select=id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy,folder,file,package,parentReference`,
      top
    );
  }

  public getRecentFiles(top: number): Promise<IGraphDriveItem[]> {
    return this.getAll<IGraphDriveItem>(`/me/drive/recent?$top=${top}`, top);
  }

  public getSharedWithMe(top: number): Promise<IGraphDriveItem[]> {
    return this.getAll<IGraphDriveItem>('/me/drive/sharedWithMe', top);
  }

  public searchMyDrive(query: string, top: number): Promise<IGraphDriveItem[]> {
    return this.getAll<IGraphDriveItem>(`/me/drive/root/search(q='${encodeURIComponent(query.replace(/'/g, "''"))}')?$top=${top}`, top);
  }

  // ---------------------------------------------------------------------------
  // Calendar, tasks, mail, teams
  // ---------------------------------------------------------------------------

  public getCalendarView(start: Date, end: Date, top: number): Promise<IGraphEvent[]> {
    const tz: string = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    return this.getAll<IGraphEvent>(
      `/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$orderby=start/dateTime&$top=${top}&$select=id,subject,webLink,isAllDay,isOnlineMeeting,onlineMeeting,start,end,location,organizer,responseStatus,showAs`,
      top,
      { headers: { Prefer: `outlook.timezone="${tz}"` } }
    );
  }

  public getTodoLists(): Promise<IGraphTodoList[]> {
    return this.getAll<IGraphTodoList>('/me/todo/lists', 50);
  }

  public getTodoTasks(listId: string, top: number): Promise<IGraphTodoTask[]> {
    return this.getAll<IGraphTodoTask>(
      `/me/todo/lists/${listId}/tasks?$filter=status ne 'completed'&$top=${top}&$orderby=createdDateTime desc`,
      top
    );
  }

  public async completeTodoTask(listId: string, taskId: string): Promise<void> {
    await this.patch(`/me/todo/lists/${listId}/tasks/${taskId}`, { status: 'completed' });
    this.cache.clear();
  }

  public async createTodoTask(listId: string, title: string): Promise<IGraphTodoTask> {
    const task: IGraphTodoTask = await this.post(`/me/todo/lists/${listId}/tasks`, { title });
    this.cache.clear();
    return task;
  }

  public getPlannerTasks(top: number): Promise<IGraphPlannerTask[]> {
    return this.getAll<IGraphPlannerTask>('/me/planner/tasks', top);
  }

  public getPlan(planId: string): Promise<{ id: string; title: string }> {
    return this.get<{ id: string; title: string }>(`/planner/plans/${planId}?$select=id,title`);
  }

  public getInboxMessages(top: number, unreadOnly: boolean): Promise<IGraphMessage[]> {
    const filter: string = unreadOnly ? '&$filter=isRead eq false' : '';
    return this.getAll<IGraphMessage>(
      `/me/mailFolders/inbox/messages?$top=${top}${filter}&$orderby=receivedDateTime desc&$select=id,subject,bodyPreview,receivedDateTime,isRead,webLink,importance,hasAttachments,from`,
      top
    );
  }

  public getInboxUnreadCount(): Promise<{ unreadItemCount: number; totalItemCount: number }> {
    return this.get('/me/mailFolders/inbox?$select=unreadItemCount,totalItemCount');
  }

  public getJoinedTeams(): Promise<IGraphTeam[]> {
    return this.getAll<IGraphTeam>('/me/joinedTeams', 100);
  }

  // ---------------------------------------------------------------------------
  // Entra applications (My Apps)
  // ---------------------------------------------------------------------------

  public getAppRoleAssignments(): Promise<IGraphAppRoleAssignment[]> {
    return this.getAll<IGraphAppRoleAssignment>('/me/appRoleAssignments', 200);
  }

  public async getServicePrincipals(ids: string[]): Promise<IGraphServicePrincipal[]> {
    if (ids.length === 0) {
      return [];
    }
    const unique: string[] = ids.filter((v, i, a) => a.indexOf(v) === i);
    const results: IGraphServicePrincipal[] = [];
    // Batch in chunks of 15 ids per $filter to keep URLs short.
    for (let i: number = 0; i < unique.length; i += 15) {
      const chunk: string[] = unique.slice(i, i + 15);
      const filter: string = chunk.map(id => `id eq '${id}'`).join(' or ');
      const page: IGraphServicePrincipal[] = await this.getAll<IGraphServicePrincipal>(
        `/servicePrincipals?$filter=${encodeURIComponent(filter)}&$select=id,appId,displayName,homepage,loginUrl,info,tags,accountEnabled`,
        chunk.length
      );
      results.push(...page);
    }
    return results;
  }
}

export function isNotFound(error: unknown): boolean {
  const e = error as { statusCode?: number; status?: number; code?: string; message?: string };
  return e?.statusCode === 404 || e?.status === 404 || e?.code === 'Request_ResourceNotFound' || e?.code === 'ResourceNotFound';
}

export function isForbidden(error: unknown): boolean {
  const e = error as { statusCode?: number; status?: number; code?: string };
  return e?.statusCode === 403 || e?.status === 403 || e?.code === 'Authorization_RequestDenied' || e?.code === 'accessDenied';
}

export function describeError(error: unknown): string {
  const e = error as { message?: string; code?: string; statusCode?: number; body?: string };
  if (!e) {
    return 'Unknown error';
  }
  if (isForbidden(e)) {
    return 'You do not have permission to view this information. An administrator may need to approve the API permissions requested by this solution.';
  }
  if (typeof e.body === 'string') {
    try {
      const parsed: { error?: { message?: string } } = JSON.parse(e.body);
      if (parsed?.error?.message) {
        return parsed.error.message;
      }
    } catch {
      /* not JSON */
    }
  }
  return e.message || e.code || 'Something went wrong';
}
