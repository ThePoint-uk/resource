import type { IDashboardState, StorageMode } from '../models';
import type { GraphService } from './GraphService';
import type { SharePointService } from './SharePointService';

export const LAYOUT_VERSION: number = 1;

export interface ILayoutStoreOptions {
  mode: StorageMode;
  listTitle: string;
  /** Instance key so different web part instances can keep distinct layouts. */
  instanceKey: string;
  userKey: string;
  graph: GraphService;
  sp: SharePointService;
}

/**
 * Persists the user's dashboard layout.
 *
 * - `local`     : browser localStorage only (always used as a fast cache / fallback).
 * - `onedrive`  : JSON file in the user's OneDrive app folder (roams across devices; needs Files.ReadWrite).
 * - `list`      : item per user in a SharePoint list (Title = user, Layout = JSON multi-line text).
 */
export class LayoutStore {
  private readonly _options: ILayoutStoreOptions;
  private _saveTimer: number | undefined;
  private _pending: IDashboardState | undefined;

  constructor(options: ILayoutStoreOptions) {
    this._options = options;
  }

  private get _localKey(): string {
    return `tpms:layout:${this._options.userKey}:${this._options.instanceKey}`;
  }

  private get _remoteFileName(): string {
    return `layout-${sanitise(this._options.instanceKey)}.json`;
  }

  public readLocal(): IDashboardState | undefined {
    try {
      const raw: string | null = window.localStorage.getItem(this._localKey);
      return raw ? (JSON.parse(raw) as IDashboardState) : undefined;
    } catch {
      return undefined;
    }
  }

  public writeLocal(state: IDashboardState): void {
    try {
      window.localStorage.setItem(this._localKey, JSON.stringify(state));
    } catch {
      /* storage unavailable */
    }
  }

  public clearLocal(): void {
    try {
      window.localStorage.removeItem(this._localKey);
    } catch {
      /* ignore */
    }
  }

  /**
   * Loads the layout: remote first (if configured), local fallback.
   */
  public async load(): Promise<{ state: IDashboardState | undefined; source: 'remote' | 'local' | 'none'; error?: string }> {
    const local: IDashboardState | undefined = this.readLocal();
    if (this._options.mode === 'local') {
      return { state: local, source: local ? 'local' : 'none' };
    }
    try {
      const remote: IDashboardState | undefined = await this._loadRemote();
      if (remote) {
        // Keep whichever copy is newer, so an offline edit is not lost.
        if (local && local.updated && remote.updated && local.updated > remote.updated) {
          return { state: local, source: 'local' };
        }
        this.writeLocal(remote);
        return { state: remote, source: 'remote' };
      }
      return { state: local, source: local ? 'local' : 'none' };
    } catch (e) {
      return { state: local, source: local ? 'local' : 'none', error: (e as Error).message };
    }
  }

  /**
   * Saves immediately to localStorage and debounces the remote write.
   */
  public save(state: IDashboardState): void {
    const stamped: IDashboardState = { ...state, version: LAYOUT_VERSION, updated: new Date().toISOString() };
    this.writeLocal(stamped);
    if (this._options.mode === 'local') {
      return;
    }
    this._pending = stamped;
    if (this._saveTimer) {
      window.clearTimeout(this._saveTimer);
    }
    this._saveTimer = window.setTimeout(() => {
      this._saveTimer = undefined;
      const toSave: IDashboardState | undefined = this._pending;
      this._pending = undefined;
      if (toSave) {
        this._saveRemote(toSave).catch((e: Error) => console.warn('[ThePoint MySite] Remote layout save failed', e));
      }
    }, 1200);
  }

  public async flush(): Promise<void> {
    if (this._saveTimer) {
      window.clearTimeout(this._saveTimer);
      this._saveTimer = undefined;
    }
    if (this._pending) {
      const toSave: IDashboardState = this._pending;
      this._pending = undefined;
      await this._saveRemote(toSave);
    }
  }

  public async reset(): Promise<void> {
    this.clearLocal();
    if (this._options.mode === 'onedrive') {
      try {
        await this._options.graph.delete(`/me/drive/special/approot:/${this._remoteFileName}`);
      } catch {
        /* nothing to delete */
      }
    } else if (this._options.mode === 'list') {
      await this._options.sp.upsertListItemByUser(this._options.listTitle, this._listKey, '');
    }
  }

  private get _listKey(): string {
    return `${this._options.userKey}|${this._options.instanceKey}`;
  }

  private async _loadRemote(): Promise<IDashboardState | undefined> {
    if (this._options.mode === 'onedrive') {
      try {
        const content: unknown = await this._options.graph.get<unknown>(`/me/drive/special/approot:/${this._remoteFileName}:/content`, { noCache: true });
        return coerce(content);
      } catch (e) {
        const err = e as { statusCode?: number; status?: number };
        if (err?.statusCode === 404 || err?.status === 404) {
          return undefined;
        }
        throw e;
      }
    }
    if (this._options.mode === 'list') {
      const item = await this._options.sp.getListItemByUser(this._options.listTitle, this._listKey);
      return item && item.Layout ? coerce(item.Layout) : undefined;
    }
    return undefined;
  }

  private async _saveRemote(state: IDashboardState): Promise<void> {
    const json: string = JSON.stringify(state);
    if (this._options.mode === 'onedrive') {
      await this._options.graph.put(`/me/drive/special/approot:/${this._remoteFileName}:/content`, json, 'application/json');
    } else if (this._options.mode === 'list') {
      await this._options.sp.upsertListItemByUser(this._options.listTitle, this._listKey, json);
    }
  }
}

function coerce(content: unknown): IDashboardState | undefined {
  try {
    const obj: IDashboardState = typeof content === 'string' ? JSON.parse(content) : (content as IDashboardState);
    if (obj && Array.isArray(obj.widgets)) {
      return obj;
    }
  } catch {
    /* corrupt */
  }
  return undefined;
}

function sanitise(value: string): string {
  return value.replace(/[^a-zA-Z0-9-_]/g, '').slice(0, 60) || 'default';
}
