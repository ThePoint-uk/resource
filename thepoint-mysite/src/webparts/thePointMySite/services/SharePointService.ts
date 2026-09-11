import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import { CacheService } from './CacheService';

export interface ISearchResultRow {
  [key: string]: string | undefined;
}

export interface ISPNewsItem {
  id: string;
  title: string;
  url: string;
  description?: string;
  imageUrl?: string;
  published?: string;
  author?: string;
  siteTitle?: string;
}

export interface ISPSite {
  id: string;
  title: string;
  url: string;
  description?: string;
  webTemplate?: string;
  logoUrl?: string;
}

export interface ISPProfileProperty {
  Key: string;
  Value: string;
}

export interface ISPUserProfile {
  DisplayName?: string;
  Email?: string;
  Title?: string;
  PictureUrl?: string;
  PersonalUrl?: string;
  UserUrl?: string;
  UserProfileProperties?: ISPProfileProperty[];
}

/**
 * SharePoint REST helper for features Graph does not (yet) cover well:
 * search-driven news, all-sites listing, user profile properties and list-based storage.
 */
export class SharePointService {
  private readonly _context: WebPartContext;
  public readonly cache: CacheService;

  constructor(context: WebPartContext, cache: CacheService) {
    this._context = context;
    this.cache = cache;
  }

  public get webUrl(): string {
    return this._context.pageContext.web.absoluteUrl;
  }

  public get siteUrl(): string {
    return this._context.pageContext.site.absoluteUrl;
  }

  public get tenantUrl(): string {
    const url: URL = new URL(this._context.pageContext.web.absoluteUrl);
    return `${url.protocol}//${url.host}`;
  }

  public get userLoginName(): string {
    return this._context.pageContext.user.loginName;
  }

  public get userEmail(): string {
    return this._context.pageContext.user.email;
  }

  private async _get<T>(url: string, cacheKey?: string): Promise<T> {
    const key: string = cacheKey || `sp:${url}`;
    const cached: T | undefined = this.cache.get<T>(key);
    if (cached !== undefined) {
      return cached;
    }
    const response: SPHttpClientResponse = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
    });
    if (!response.ok) {
      throw new Error(`SharePoint request failed (${response.status}): ${await safeText(response)}`);
    }
    const json: T = await response.json();
    this.cache.put(key, json);
    return json;
  }

  private async _post<T>(url: string, body: unknown, extraHeaders?: Record<string, string>): Promise<T> {
    const response: SPHttpClientResponse = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        ...(extraHeaders || {})
      },
      body: JSON.stringify(body)
    });
    if (!response.ok) {
      throw new Error(`SharePoint request failed (${response.status}): ${await safeText(response)}`);
    }
    if (response.status === 204) {
      return undefined as unknown as T;
    }
    return response.json();
  }

  // ---------------------------------------------------------------------------
  // Search
  // ---------------------------------------------------------------------------

  public async search(queryText: string, selectProperties: string[], rowLimit: number, sortList?: string, refinementFilters?: string): Promise<ISearchResultRow[]> {
    const params: string[] = [
      `querytext='${encodeURIComponent(queryText.replace(/'/g, "''"))}'`,
      `selectproperties='${encodeURIComponent(selectProperties.join(','))}'`,
      `rowlimit=${rowLimit}`,
      `trimduplicates=false`,
      `clienttype='ContentSearchRegular'`
    ];
    if (sortList) {
      params.push(`sortlist='${encodeURIComponent(sortList)}'`);
    }
    if (refinementFilters) {
      params.push(`refinementfilters='${encodeURIComponent(refinementFilters)}'`);
    }
    const url: string = `${this.webUrl}/_api/search/query?${params.join('&')}`;
    const json: {
      PrimaryQueryResult?: { RelevantResults?: { Table?: { Rows?: { Cells: { Key: string; Value: string }[] }[] } } };
    } = await this._get(url);
    const rows = json.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.map(row => {
      const obj: ISearchResultRow = {};
      row.Cells.forEach(cell => {
        obj[cell.Key] = cell.Value;
      });
      return obj;
    });
  }

  public async getNews(rowLimit: number, scope: 'tenant' | 'site' | 'hub'): Promise<ISPNewsItem[]> {
    let query: string = 'PromotedState:2 AND FileExtension:aspx AND IsDocument:1 -ContentTypeId:0x0120*';
    if (scope === 'site') {
      query += ` AND Path:"${this.siteUrl}"`;
    } else if (scope === 'hub') {
      const hubId: string | undefined = this._context.pageContext.legacyPageContext?.hubSiteId;
      if (hubId) {
        query += ` AND DepartmentId:{${hubId}}`;
      }
    }
    const rows: ISearchResultRow[] = await this.search(
      query,
      ['Title', 'Path', 'Description', 'PictureThumbnailURL', 'FirstPublishedDate', 'Author', 'SiteTitle', 'UniqueId', 'BannerImageUrl'],
      rowLimit,
      'FirstPublishedDate:descending'
    );
    return rows.map(r => ({
      id: r.UniqueId || r.Path || '',
      title: r.Title || '',
      url: r.Path || '',
      description: r.Description,
      imageUrl: r.PictureThumbnailURL || r.BannerImageUrl,
      published: r.FirstPublishedDate,
      author: (r.Author || '').split(';')[0],
      siteTitle: r.SiteTitle
    }));
  }

  public async getSites(rowLimit: number, query?: string): Promise<ISPSite[]> {
    let queryText: string = 'contentclass:STS_Site -WebTemplate:SPSPERS -WebTemplate:APPCATALOG -WebTemplate:SRCHCEN -WebTemplate:REDIRECTSITE';
    if (query) {
      queryText = `${query}* AND ${queryText}`;
    }
    const rows: ISearchResultRow[] = await this.search(
      queryText,
      ['Title', 'Path', 'Description', 'SiteLogo', 'WebTemplate', 'SiteId', 'LastModifiedTime'],
      rowLimit,
      query ? undefined : 'LastModifiedTime:descending'
    );
    return rows.map(r => ({
      id: r.SiteId || r.Path || '',
      title: r.Title || '',
      url: r.Path || '',
      description: r.Description,
      webTemplate: r.WebTemplate,
      logoUrl: r.SiteLogo
    }));
  }

  // ---------------------------------------------------------------------------
  // User profile (SharePoint profile service)
  // ---------------------------------------------------------------------------

  public getMyProfile(): Promise<ISPUserProfile> {
    return this._get<ISPUserProfile>(`${this.webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, 'sp:myprofile');
  }

  public async setMyProfileProperty(propertyName: string, value: string): Promise<void> {
    await this._post(`${this.webUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`, {
      accountName: this._context.pageContext.user.loginName.indexOf('i:0#.f|membership|') === 0
        ? this._context.pageContext.user.loginName
        : `i:0#.f|membership|${this._context.pageContext.user.loginName}`,
      propertyName,
      propertyValue: value
    });
    this.cache.remove('sp:myprofile');
  }

  // ---------------------------------------------------------------------------
  // List-based layout storage
  // ---------------------------------------------------------------------------

  public async getListItemByUser(listTitle: string, userKey: string): Promise<{ Id: number; Layout: string } | undefined> {
    const url: string = `${this.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle.replace(/'/g, "''"))}')/items?$select=Id,Layout&$filter=Title eq '${encodeURIComponent(userKey.replace(/'/g, "''"))}'&$top=1`;
    const response: SPHttpClientResponse = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
    });
    if (!response.ok) {
      throw new Error(`List read failed (${response.status})`);
    }
    const json: { value: { Id: number; Layout: string }[] } = await response.json();
    return json.value[0];
  }

  public async upsertListItemByUser(listTitle: string, userKey: string, layout: string): Promise<void> {
    const existing = await this.getListItemByUser(listTitle, userKey);
    const listPath: string = `${this.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitle.replace(/'/g, "''"))}')`;
    if (existing) {
      await this._post(`${listPath}/items(${existing.Id})`, { Layout: layout }, { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' });
    } else {
      await this._post(`${listPath}/items`, { Title: userKey, Layout: layout });
    }
  }
}

async function safeText(response: SPHttpClientResponse): Promise<string> {
  try {
    return (await response.text()).slice(0, 300);
  } catch {
    return '';
  }
}
