import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import {
  IFileResult, IMailResult, IEventResult, ISiteResult, IRefiner, IRefinerBucket,
  ISearchPage, ISuggestions, IPersonResult, SearchScope, SortMode
} from '../models';

/**
 * All search runs through the Microsoft Graph /search/query endpoint with the
 * signed-in user's delegated token, so every result set is security-trimmed to
 * what that user can already see. Nothing here elevates permissions.
 */

export interface IFileSearchOptions {
  from: number;
  size: number;
  sort: SortMode;
  aggregationFilters: string[];
  /** Extra KQL appended to the user query (scope restriction, policies template…). */
  extraKql?: string;
  withRefiners: boolean;
  enableSpelling: boolean;
}

interface IGraphHit {
  hitId?: string;
  summary?: string;
  resource?: {
    [key: string]: unknown;
  };
}

interface IGraphHitsContainer {
  hits?: IGraphHit[];
  total?: number;
  moreResultsAvailable?: boolean;
  aggregations?: IGraphAggregation[];
}

interface IGraphAggregation {
  field?: string;
  buckets?: {
    key?: string;
    count?: number;
    aggregationFilterToken?: string;
  }[];
}

interface IGraphSearchResponse {
  value?: {
    hitsContainers?: IGraphHitsContainer[];
    queryAlterationResponse?: {
      queryAlteration?: {
        alteredQueryString?: string;
      };
    };
  }[];
}

export class GraphSearchService {

  private _clientPromise: Promise<MSGraphClientV3>;
  private _siteUrl: string;
  private _hubSiteId?: string;

  constructor(factory: MSGraphClientFactory, siteUrl: string, hubSiteId?: string) {
    this._clientPromise = factory.getClient('3');
    this._siteUrl = siteUrl;
    this._hubSiteId = hubSiteId;
  }

  public scopeKql(scope: SearchScope): string | undefined {
    if (scope === 'site') {
      return `path:"${this._siteUrl}"`;
    }
    if (scope === 'hub' && this._hubSiteId) {
      return `DepartmentId:{${this._hubSiteId}}`;
    }
    return undefined;
  }

  /* ------------------------------------------------------------------ */
  /* Files (driveItem: SharePoint document libraries + OneDrive)         */
  /* ------------------------------------------------------------------ */

  public async searchFiles(query: string, options: IFileSearchOptions): Promise<ISearchPage<IFileResult>> {
    const client = await this._clientPromise;
    const queryString = options.extraKql ? `(${query}) ${options.extraKql}` : query;

    const request: { [key: string]: unknown } = {
      entityTypes: ['driveItem'],
      query: { queryString },
      from: options.from,
      size: options.size,
      fields: [
        'name', 'webUrl', 'lastModifiedDateTime', 'lastModifiedBy', 'createdBy',
        'parentReference', 'size', 'fileType', 'id', 'siteName', 'title'
      ]
    };

    if (options.sort !== 'relevance') {
      request.sortProperties = [{
        name: 'lastModifiedDateTime',
        isDescending: options.sort === 'newest'
      }];
    }
    if (options.withRefiners) {
      request.aggregations = [
        { field: 'fileType', size: 8, bucketDefinition: { sortBy: 'count', isDescending: true, minimumCount: 0 } },
        {
          field: 'lastModifiedTime',
          bucketDefinition: {
            sortBy: 'keyAsString', isDescending: true, minimumCount: 0,
            ranges: this._dateRanges()
          }
        }
      ];
    }
    if (options.aggregationFilters.length > 0) {
      request.aggregationFilters = options.aggregationFilters;
    }

    const body: { [key: string]: unknown } = { requests: [request] };
    if (options.enableSpelling) {
      (body.requests as unknown[])[0] = {
        ...(request),
        queryAlterationOptions: { enableSuggestion: true, enableModification: false }
      };
    }

    const response: IGraphSearchResponse = await client.api('/search/query').post(body);
    const container = this._container(response);
    const items = (container.hits || []).map(hit => this._toFile(hit));

    return {
      items,
      total: container.total || items.length,
      moreAvailable: !!container.moreResultsAvailable,
      refiners: this._toRefiners(container.aggregations),
      spellingSuggestion: this._spelling(response)
    };
  }

  private _toFile(hit: IGraphHit): IFileResult {
    const r = (hit.resource || {}) as {
      name?: string; webUrl?: string; lastModifiedDateTime?: string; size?: number; id?: string;
      lastModifiedBy?: { user?: { displayName?: string } };
      createdBy?: { user?: { displayName?: string } };
      parentReference?: { driveId?: string; siteId?: string; path?: string };
      listItem?: { fields?: { [key: string]: unknown } };
    };
    const name = r.name || '';
    const dot = name.lastIndexOf('.');
    const webUrl = r.webUrl || '';
    const isOneDrive = webUrl.indexOf('-my.sharepoint.com') !== -1;
    return {
      key: hit.hitId || r.id || webUrl,
      name: dot > 0 ? name.substring(0, dot) : name,
      fileType: dot > 0 ? name.substring(dot + 1).toLowerCase() : '',
      webUrl,
      driveId: r.parentReference ? r.parentReference.driveId : undefined,
      itemId: r.id,
      siteName: isOneDrive ? 'OneDrive' : this._siteNameFromUrl(webUrl),
      modified: r.lastModifiedDateTime,
      modifiedBy: r.lastModifiedBy && r.lastModifiedBy.user ? r.lastModifiedBy.user.displayName : undefined,
      author: r.createdBy && r.createdBy.user ? r.createdBy.user.displayName : undefined,
      size: r.size,
      summary: hit.summary,
      isOneDrive
    };
  }

  private _siteNameFromUrl(webUrl: string): string {
    const match = /\/(?:sites|teams)\/([^/]+)/i.exec(webUrl);
    return match ? decodeURIComponent(match[1]) : 'SharePoint';
  }

  private _dateRanges(): { to?: string; from?: string }[] {
    const now = new Date();
    const iso = (daysAgo: number): string => {
      const d = new Date(now.getTime() - daysAgo * 86400000);
      return d.toISOString();
    };
    // Buckets surface as "older than a year / within a year / month / week / day".
    return [
      { to: iso(365) },
      { from: iso(365), to: iso(30) },
      { from: iso(30), to: iso(7) },
      { from: iso(7), to: iso(1) },
      { from: iso(1) }
    ];
  }

  /* ------------------------------------------------------------------ */
  /* Mail (message)                                                      */
  /* ------------------------------------------------------------------ */

  public async searchMail(
    query: string, from: number, size: number, extraKql?: string
  ): Promise<ISearchPage<IMailResult>> {
    const client = await this._clientPromise;
    const queryString = extraKql ? `(${query}) ${extraKql}` : query;
    const response: IGraphSearchResponse = await client.api('/search/query').post({
      requests: [{
        entityTypes: ['message'],
        query: { queryString },
        from,
        size
      }]
    });
    const container = this._container(response);
    const items = (container.hits || []).map(hit => {
      const r = (hit.resource || {}) as {
        id?: string; subject?: string; bodyPreview?: string; receivedDateTime?: string;
        hasAttachments?: boolean; webLink?: string;
        from?: { emailAddress?: { name?: string; address?: string } };
      };
      return {
        key: hit.hitId || r.id || Math.random().toString(36),
        subject: r.subject || '(no subject)',
        from: r.from && r.from.emailAddress ? (r.from.emailAddress.name || r.from.emailAddress.address || '') : '',
        fromEmail: r.from && r.from.emailAddress ? r.from.emailAddress.address : undefined,
        preview: hit.summary || r.bodyPreview,
        received: r.receivedDateTime,
        hasAttachments: !!r.hasAttachments,
        webLink: r.webLink
      } as IMailResult;
    });
    return { items, total: container.total || items.length, moreAvailable: !!container.moreResultsAvailable };
  }

  /* ------------------------------------------------------------------ */
  /* Calendar (event)                                                    */
  /* ------------------------------------------------------------------ */

  public async searchEvents(query: string, from: number, size: number): Promise<ISearchPage<IEventResult>> {
    const client = await this._clientPromise;
    const response: IGraphSearchResponse = await client.api('/search/query').post({
      requests: [{
        entityTypes: ['event'],
        query: { queryString: query },
        from,
        size
      }]
    });
    const container = this._container(response);
    const items = (container.hits || []).map(hit => this._toEvent(hit));
    return { items, total: container.total || items.length, moreAvailable: !!container.moreResultsAvailable };
  }

  private _toEvent(hit: IGraphHit): IEventResult {
    const r = (hit.resource || {}) as {
      id?: string; subject?: string; webLink?: string; isOnlineMeeting?: boolean;
      start?: { dateTime?: string }; end?: { dateTime?: string };
      location?: { displayName?: string };
      organizer?: { emailAddress?: { name?: string } };
    };
    return {
      key: hit.hitId || r.id || Math.random().toString(36),
      subject: r.subject || '(no subject)',
      start: r.start ? r.start.dateTime : undefined,
      end: r.end ? r.end.dateTime : undefined,
      location: r.location ? r.location.displayName : undefined,
      organizer: r.organizer && r.organizer.emailAddress ? r.organizer.emailAddress.name : undefined,
      isOnline: !!r.isOnlineMeeting,
      webLink: r.webLink
    };
  }

  /* ------------------------------------------------------------------ */
  /* Sites (site)                                                        */
  /* ------------------------------------------------------------------ */

  public async searchSites(query: string, from: number, size: number): Promise<ISearchPage<ISiteResult>> {
    const client = await this._clientPromise;
    const response: IGraphSearchResponse = await client.api('/search/query').post({
      requests: [{
        entityTypes: ['site'],
        query: { queryString: query },
        from,
        size
      }]
    });
    const container = this._container(response);
    const items = (container.hits || []).map(hit => {
      const r = (hit.resource || {}) as { id?: string; displayName?: string; name?: string; webUrl?: string; description?: string };
      return {
        key: hit.hitId || r.id || r.webUrl || Math.random().toString(36),
        name: r.displayName || r.name || '',
        webUrl: r.webUrl || '',
        description: r.description || hit.summary
      } as ISiteResult;
    });
    return { items, total: container.total || items.length, moreAvailable: !!container.moreResultsAvailable };
  }

  /* ------------------------------------------------------------------ */
  /* Suggestions (as-you-type)                                           */
  /* ------------------------------------------------------------------ */

  public async suggest(query: string, peopleSearch: (q: string, top: number) => Promise<IPersonResult[]>): Promise<ISuggestions> {
    const [filesPage, people] = await Promise.all([
      this.searchFiles(query, {
        from: 0, size: 3, sort: 'relevance', aggregationFilters: [], withRefiners: false, enableSpelling: false
      }).catch((): ISearchPage<IFileResult> => ({ items: [], total: 0, moreAvailable: false })),
      peopleSearch(query, 2).catch((): IPersonResult[] => [])
    ]);
    return { files: filesPage.items, people };
  }

  /* ------------------------------------------------------------------ */
  /* Shared plumbing                                                     */
  /* ------------------------------------------------------------------ */

  private _container(response: IGraphSearchResponse): IGraphHitsContainer {
    if (response.value && response.value.length > 0) {
      const containers = response.value[0].hitsContainers;
      if (containers && containers.length > 0) {
        return containers[0];
      }
    }
    return {};
  }

  private _spelling(response: IGraphSearchResponse): string | undefined {
    if (response.value && response.value.length > 0) {
      const alteration = response.value[0].queryAlterationResponse;
      if (alteration && alteration.queryAlteration) {
        return alteration.queryAlteration.alteredQueryString;
      }
    }
    return undefined;
  }

  private _toRefiners(aggregations?: IGraphAggregation[]): IRefiner[] | undefined {
    if (!aggregations || aggregations.length === 0) {
      return undefined;
    }
    const labels: { [field: string]: string } = {
      fileType: 'File type',
      lastModifiedTime: 'Last modified'
    };
    const dateLabels = ['Older than a year', 'Past year', 'Past month', 'Past week', 'Past 24 hours'];
    return aggregations.map(aggregation => {
      const field = aggregation.field || '';
      const buckets: IRefinerBucket[] = (aggregation.buckets || [])
        .filter(bucket => !!bucket.aggregationFilterToken && (bucket.count || 0) > 0)
        .map((bucket, index) => ({
          key: bucket.key || String(index),
          label: field === 'lastModifiedTime'
            ? (dateLabels[index] !== undefined ? dateLabels[index] : bucket.key || '')
            : (bucket.key || ''),
          count: bucket.count || 0,
          filterToken: `${field}:${bucket.aggregationFilterToken}`
        }));
      return { field, label: labels[field] || field, buckets };
    }).filter(refiner => refiner.buckets.length > 0);
  }
}
