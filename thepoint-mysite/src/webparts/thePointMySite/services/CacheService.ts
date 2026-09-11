/**
 * Small TTL cache backed by sessionStorage with an in-memory fallback.
 * Keys are namespaced per user so shared machines never leak data across accounts.
 */
export class CacheService {
  private readonly _prefix: string;
  private readonly _memory: Map<string, { expires: number; value: unknown }> = new Map();
  private _ttlMs: number;

  constructor(namespace: string, ttlMinutes: number) {
    this._prefix = `tpms:${namespace}:`;
    this._ttlMs = Math.max(0, ttlMinutes) * 60 * 1000;
  }

  public get ttlMinutes(): number {
    return this._ttlMs / 60000;
  }

  public set ttlMinutes(value: number) {
    this._ttlMs = Math.max(0, value) * 60 * 1000;
  }

  public get<T>(key: string): T | undefined {
    if (this._ttlMs === 0) {
      return undefined;
    }
    const fullKey: string = this._prefix + key;
    const now: number = Date.now();
    const mem = this._memory.get(fullKey);
    if (mem && mem.expires > now) {
      return mem.value as T;
    }
    try {
      const raw: string | null = window.sessionStorage.getItem(fullKey);
      if (raw) {
        const parsed: { expires: number; value: T } = JSON.parse(raw);
        if (parsed.expires > now) {
          this._memory.set(fullKey, parsed);
          return parsed.value;
        }
        window.sessionStorage.removeItem(fullKey);
      }
    } catch {
      /* storage unavailable */
    }
    return undefined;
  }

  public put<T>(key: string, value: T): void {
    if (this._ttlMs === 0) {
      return;
    }
    const fullKey: string = this._prefix + key;
    const entry = { expires: Date.now() + this._ttlMs, value };
    this._memory.set(fullKey, entry);
    try {
      window.sessionStorage.setItem(fullKey, JSON.stringify(entry));
    } catch {
      /* quota or storage unavailable */
    }
  }

  public remove(key: string): void {
    const fullKey: string = this._prefix + key;
    this._memory.delete(fullKey);
    try {
      window.sessionStorage.removeItem(fullKey);
    } catch {
      /* ignore */
    }
  }

  public clear(): void {
    this._memory.clear();
    try {
      const keys: string[] = [];
      for (let i: number = 0; i < window.sessionStorage.length; i++) {
        const k: string | null = window.sessionStorage.key(i);
        if (k && k.indexOf(this._prefix) === 0) {
          keys.push(k);
        }
      }
      keys.forEach(k => window.sessionStorage.removeItem(k));
    } catch {
      /* ignore */
    }
  }

  /**
   * Returns the cached value when present, otherwise executes the factory and caches the result.
   */
  public async getOrFetch<T>(key: string, factory: () => Promise<T>): Promise<T> {
    const cached: T | undefined = this.get<T>(key);
    if (cached !== undefined) {
      return cached;
    }
    const value: T = await factory();
    this.put(key, value);
    return value;
  }
}
