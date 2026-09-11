import type { GraphService, IGraphGroup } from './GraphService';

export const AUDIENCE_EVERYONE: string = '@everyone';
export const AUDIENCE_MANAGERS: string = '@managers';

/**
 * Resolves owner-defined audiences ("Finance, Sales Leads, @managers") against the
 * signed-in user's group membership. Group names, object ids and mail addresses are accepted.
 * Membership is fetched once per session and cached.
 */
export class AudienceService {
  private readonly _graph: GraphService;
  private _groups: Promise<IGraphGroup[]> | undefined;
  private _isManager: Promise<boolean> | undefined;

  constructor(graph: GraphService) {
    this._graph = graph;
  }

  public static parse(audiences: string | undefined): string[] {
    return (audiences || '')
      .split(/[,;\n]/)
      .map(a => a.trim())
      .filter(a => a.length > 0);
  }

  /** True when the audience string is blank or contains @everyone. */
  public static isOpen(audiences: string | undefined): boolean {
    const list: string[] = AudienceService.parse(audiences);
    return list.length === 0 || list.some(a => a.toLowerCase() === AUDIENCE_EVERYONE);
  }

  private _getGroups(): Promise<IGraphGroup[]> {
    if (!this._groups) {
      this._groups = this._graph.getMyGroups().catch((e: Error) => {
        console.warn('[ThePoint MySite] Could not read group membership; audience-targeted widgets will be hidden.', e);
        return [];
      });
    }
    return this._groups;
  }

  private _getIsManager(): Promise<boolean> {
    if (!this._isManager) {
      this._isManager = this._graph.hasDirectReports().catch(() => false);
    }
    return this._isManager;
  }

  /**
   * Returns true when the user is in at least one of the listed audiences.
   * A blank list means everyone.
   */
  public async isMember(audiences: string | undefined): Promise<boolean> {
    if (AudienceService.isOpen(audiences)) {
      return true;
    }
    const wanted: string[] = AudienceService.parse(audiences).map(a => a.toLowerCase());
    if (wanted.indexOf(AUDIENCE_MANAGERS) >= 0 && (await this._getIsManager())) {
      return true;
    }
    const groups: IGraphGroup[] = await this._getGroups();
    return groups.some(g =>
      wanted.indexOf(g.id.toLowerCase()) >= 0 ||
      wanted.indexOf((g.displayName || '').toLowerCase()) >= 0 ||
      (!!g.mail && wanted.indexOf(g.mail.toLowerCase()) >= 0) ||
      (!!g.mailNickname && wanted.indexOf(g.mailNickname.toLowerCase()) >= 0)
    );
  }

  /** Evaluates many audience strings at once, returning a map of audience string -> membership. */
  public async evaluate(audienceStrings: string[]): Promise<Map<string, boolean>> {
    const result: Map<string, boolean> = new Map();
    const unique: string[] = audienceStrings.filter((v, i, a) => a.indexOf(v) === i);
    await Promise.all(
      unique.map(async a => {
        result.set(a, await this.isMember(a));
      })
    );
    return result;
  }
}
