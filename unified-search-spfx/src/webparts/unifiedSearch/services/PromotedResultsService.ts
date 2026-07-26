import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPromotedResult } from '../models';

/**
 * Promoted results ("bookmarks") driven by an ordinary SharePoint list, so
 * content owners manage them without touching the admin centre or redeploying.
 *
 * Expected list columns (internal names):
 *   Title      (single line)  - card headline
 *   Keywords   (multi line)   - trigger phrases, one per line or ;-separated
 *   Description(multi line)   - card body
 *   LinkUrl    (hyperlink or single line) - where the button goes
 *   ButtonText (single line)  - button label, defaults to "Open"
 *   Enabled    (yes/no)       - kill switch
 *   StartDate  (date, optional) - publish window start
 *   EndDate    (date, optional) - publish window end
 *
 * The provisioning script in /provisioning creates this list ready-made.
 */

interface IPromotedListItem {
  title: string;
  keywords: string[];
  description?: string;
  linkUrl?: string;
  buttonText?: string;
  enabled: boolean;
  startDate?: Date;
  endDate?: Date;
}

interface IRawListItem {
  Title?: string;
  Keywords?: string;
  Description?: string;
  LinkUrl?: string | { Url?: string };
  ButtonText?: string;
  Enabled?: boolean;
  StartDate?: string;
  EndDate?: string;
}

const CACHE_TTL_MS = 5 * 60 * 1000;

export class PromotedResultsService {

  private _spHttpClient: SPHttpClient;
  private _defaultWebUrl: string;
  private _cache: { fetchedAt: number; cacheKey: string; items: IPromotedListItem[] } | undefined;

  constructor(spHttpClient: SPHttpClient, defaultWebUrl: string) {
    this._spHttpClient = spHttpClient;
    this._defaultWebUrl = defaultWebUrl;
  }

  /** Promoted results whose keywords match the query, best match first. */
  public async getMatches(query: string, listTitle: string, webUrl: string): Promise<IPromotedResult[]> {
    const items = await this._getItems(listTitle, webUrl);
    const normalizedQuery = query.trim().toLowerCase();
    if (!normalizedQuery) { return []; }
    const now = new Date();

    return items
      .filter(item => item.enabled)
      .filter(item => !item.startDate || item.startDate <= now)
      .filter(item => !item.endDate || item.endDate >= now)
      .map(item => ({ item, score: this._matchScore(item, normalizedQuery) }))
      .filter(entry => entry.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 2)
      .map(entry => ({
        title: entry.item.title,
        description: entry.item.description,
        linkUrl: entry.item.linkUrl,
        buttonText: entry.item.buttonText
      }));
  }

  /**
   * A keyword phrase matches when the whole phrase appears in the query, or
   * the query appears in the phrase ("annual" finds "annual leave"). Exact
   * phrase hits outrank partial ones.
   */
  private _matchScore(item: IPromotedListItem, normalizedQuery: string): number {
    let score = 0;
    for (const keyword of item.keywords) {
      if (keyword === normalizedQuery) { score = Math.max(score, 3); }
      else if (normalizedQuery.indexOf(keyword) !== -1) { score = Math.max(score, 2); }
      else if (keyword.indexOf(normalizedQuery) !== -1) { score = Math.max(score, 1); }
    }
    return score;
  }

  private async _getItems(listTitle: string, webUrl: string): Promise<IPromotedListItem[]> {
    const web = (webUrl || this._defaultWebUrl).replace(/\/+$/, '');
    const cacheKey = `${web}|${listTitle}`;
    if (this._cache && this._cache.cacheKey === cacheKey &&
        Date.now() - this._cache.fetchedAt < CACHE_TTL_MS) {
      return this._cache.items;
    }

    let items: IPromotedListItem[] = [];
    try {
      const escapedTitle = listTitle.replace(/'/g, "''");
      const url = `${web}/_api/web/lists/getbytitle('${encodeURIComponent(escapedTitle)}')/items` +
        `?$select=Title,Keywords,Description,LinkUrl,ButtonText,Enabled,StartDate,EndDate&$top=500`;
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const payload: { value?: IRawListItem[] } = await response.json();
        items = (payload.value || []).map(raw => this._toItem(raw)).filter(item => !!item.title);
      }
    } catch {
      // Missing list or no access: promoted results are optional, so fail silent,
      // and the cache stops us re-querying on every keystroke.
    }

    this._cache = { fetchedAt: Date.now(), cacheKey, items };
    return items;
  }

  private _toItem(raw: IRawListItem): IPromotedListItem {
    const linkRaw = raw.LinkUrl;
    const linkUrl = typeof linkRaw === 'string'
      ? linkRaw
      : (linkRaw && linkRaw.Url ? linkRaw.Url : undefined);
    return {
      title: raw.Title || '',
      keywords: (raw.Keywords || '')
        .split(/[;\n,]/)
        .map(keyword => keyword.trim().toLowerCase())
        .filter(keyword => keyword.length > 0),
      description: raw.Description,
      linkUrl,
      buttonText: raw.ButtonText,
      // A missing Enabled column (or null value) counts as enabled.
      enabled: raw.Enabled !== false,
      startDate: raw.StartDate ? new Date(raw.StartDate) : undefined,
      endDate: raw.EndDate ? new Date(raw.EndDate) : undefined
    };
  }
}
