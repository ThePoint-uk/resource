import * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { IAudienceLayout, IDashboardState, IWidgetConfig, IWidgetDefinition, IWidgetInstance, IThePointMySiteWebPartProps, WidgetSize } from '../models';
import { LayoutStore, LAYOUT_VERSION } from '../services/LayoutStore';
import type { GraphService } from '../services/GraphService';
import type { SharePointService } from '../services/SharePointService';
import { AudienceService } from '../services/AudienceService';
import { isForbidden } from '../services/GraphService';
import { getWidgetDefinition, getWidgetDefinitions } from '../widgets/registry';
import { uid } from '../utils/format';

export interface IDashboardContextValue {
  context: WebPartContext;
  graph: GraphService;
  sp: SharePointService;
  theme: IReadonlyTheme | undefined;
  isDarkTheme: boolean;
  props: IThePointMySiteWebPartProps;
  state: IDashboardState;
  ready: boolean;
  loadSource: 'remote' | 'local' | 'none';
  loadError: string | undefined;
  editMode: boolean;
  canCustomise: boolean;
  isPageEditMode: boolean;
  /** Widgets the current user may add: enabled, in audience and (optionally) permission-probed. */
  availableWidgets: IWidgetDefinition[];
  /** Widget types hidden because their Graph permission has not been consented. */
  unavailableTypes: Set<string>;
  isPinned: (type: string) => boolean;
  lockedKeys: (type: string) => string[];
  setEditMode: (value: boolean) => void;
  addWidget: (type: string, afterId?: string) => void;
  removeWidget: (id: string) => void;
  moveWidget: (id: string, toIndex: number) => void;
  updateWidget: (id: string, patch: Partial<IWidgetInstance>) => void;
  updateWidgetSettings: (id: string, patch: Record<string, unknown>) => void;
  resetLayout: () => Promise<void>;
  exportLayout: () => string;
  importLayout: (json: string) => boolean;
  saveAsDefault: ((audience: string) => void) | undefined;
  refreshAll: () => void;
  refreshToken: number;
}

const DashboardContext: React.Context<IDashboardContextValue | undefined> = React.createContext<IDashboardContextValue | undefined>(undefined);

export function useDashboard(): IDashboardContextValue {
  const value: IDashboardContextValue | undefined = React.useContext(DashboardContext);
  if (!value) {
    throw new Error('useDashboard must be used within DashboardProvider');
  }
  return value;
}

export interface IDashboardProviderProps {
  context: WebPartContext;
  graph: GraphService;
  sp: SharePointService;
  theme: IReadonlyTheme | undefined;
  isDarkTheme: boolean;
  props: IThePointMySiteWebPartProps;
  isPageEditMode: boolean;
  instanceId: string;
  onSaveAsDefault?: (json: string, audience: string) => void;
  children?: React.ReactNode;
}

// -----------------------------------------------------------------------------
// Owner rules helpers (pure, reused by the property pane)
// -----------------------------------------------------------------------------

export function widgetConfig(props: IThePointMySiteWebPartProps, type: string): IWidgetConfig {
  const cfg: IWidgetConfig = { ...(props.widgetConfig?.[type] || {}) };
  if (props.enabledWidgets && props.enabledWidgets[type] === false && cfg.enabled === undefined) {
    cfg.enabled = false;
  }
  return cfg;
}

export function isEnabled(props: IThePointMySiteWebPartProps, type: string): boolean {
  return widgetConfig(props, type).enabled !== false;
}

export function schemaDefaults(def: IWidgetDefinition): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  (def.settingsSchema || []).forEach(field => {
    if (field.defaultValue !== undefined) {
      result[field.key] = field.defaultValue;
    }
  });
  return result;
}

/** Schema defaults overridden by the owner's default settings JSON. */
export function ownerDefaults(props: IThePointMySiteWebPartProps, def: IWidgetDefinition): Record<string, unknown> {
  const base: Record<string, unknown> = schemaDefaults(def);
  const raw: string | undefined = widgetConfig(props, def.type).settings;
  if (raw && raw.trim()) {
    try {
      const parsed: Record<string, unknown> = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        return { ...base, ...parsed };
      }
    } catch {
      /* ignore invalid owner JSON */
    }
  }
  return base;
}

export function lockedSettingKeys(props: IThePointMySiteWebPartProps, type: string): string[] {
  const locked: Record<string, boolean> = widgetConfig(props, type).locked || {};
  return Object.keys(locked).filter(k => locked[k]);
}

/** Types the current user is allowed to have on the dashboard (enabled + audience). */
export function computeAllowedTypes(props: IThePointMySiteWebPartProps, audienceMap: Map<string, boolean> | undefined): Set<string> {
  const allowed: Set<string> = new Set();
  getWidgetDefinitions().forEach(def => {
    const cfg: IWidgetConfig = widgetConfig(props, def.type);
    if (cfg.enabled === false) {
      return;
    }
    if (AudienceService.isOpen(cfg.audiences)) {
      allowed.add(def.type);
      return;
    }
    if (audienceMap && audienceMap.get(cfg.audiences || '') === true) {
      allowed.add(def.type);
    }
  });
  return allowed;
}

export function parseLayout(json: string | undefined): IDashboardState | undefined {
  if (!json || !json.trim()) {
    return undefined;
  }
  try {
    const parsed: IDashboardState = JSON.parse(json);
    if (parsed && Array.isArray(parsed.widgets)) {
      return parsed;
    }
  } catch {
    /* invalid */
  }
  return undefined;
}

export function parseAudienceLayouts(json: string | undefined): IAudienceLayout[] {
  if (!json || !json.trim()) {
    return [];
  }
  try {
    const parsed: IAudienceLayout[] = JSON.parse(json);
    return Array.isArray(parsed) ? parsed.filter(e => e && e.layout && Array.isArray(e.layout.widgets)) : [];
  } catch {
    return [];
  }
}

const STARTER: { type: string; size: WidgetSize }[] = [
  { type: 'profile', size: 1 },
  { type: 'calendar', size: 1 },
  { type: 'tasks', size: 1 },
  { type: 'recentDocs', size: 2 },
  { type: 'quickLinks', size: 1 },
  { type: 'team', size: 1 },
  { type: 'sites', size: 1 },
  { type: 'onedrive', size: 1 },
  { type: 'news', size: 2 },
  { type: 'apps', size: 1 }
];

/**
 * Picks the default layout for this user: first matching audience layout, else the
 * page default, else the built-in starter set.
 */
export function resolveDefaultLayout(props: IThePointMySiteWebPartProps, audienceMap: Map<string, boolean> | undefined): IDashboardState {
  for (const entry of parseAudienceLayouts(props.defaultLayouts)) {
    if (AudienceService.isOpen(entry.audiences) || (audienceMap && audienceMap.get(entry.audiences) === true)) {
      return entry.layout;
    }
  }
  const page: IDashboardState | undefined = parseLayout(props.defaultLayout);
  if (page) {
    return page;
  }
  return { version: LAYOUT_VERSION, widgets: STARTER.map(s => ({ id: uid(s.type), type: s.type, size: s.size, settings: {} })) };
}

/**
 * Applies owner rules to a layout: drops disallowed widgets, clamps sizes, fills settings
 * from owner defaults, enforces locked settings and inserts missing pinned widgets.
 */
export function normalise(state: IDashboardState, props: IThePointMySiteWebPartProps, columns: number, allowed: Set<string>, defaultLayout?: IDashboardState): IDashboardState {
  const seen: Set<string> = new Set();
  const widgets: IWidgetInstance[] = [];
  const apply = (w: IWidgetInstance, def: IWidgetDefinition): IWidgetInstance => {
    let id: string = w.id || uid(w.type);
    if (seen.has(id)) {
      id = uid(w.type);
    }
    seen.add(id);
    const defaults: Record<string, unknown> = ownerDefaults(props, def);
    const settings: Record<string, unknown> = { ...defaults, ...(w.settings || {}) };
    lockedSettingKeys(props, def.type).forEach(k => {
      if (defaults[k] !== undefined) {
        settings[k] = defaults[k];
      }
    });
    const size: WidgetSize = Math.max(1, Math.min(columns, Number(w.size) || def.defaultSize)) as WidgetSize;
    return { id, type: w.type, size, collapsed: !!w.collapsed, title: w.title, settings };
  };
  for (const w of state.widgets) {
    const def: IWidgetDefinition | undefined = getWidgetDefinition(w.type);
    if (!def || !allowed.has(w.type)) {
      continue;
    }
    widgets.push(apply(w, def));
  }
  // Pinned widgets must be present. Insert at their default-layout position where known.
  getWidgetDefinitions().forEach(def => {
    const cfg: IWidgetConfig = widgetConfig(props, def.type);
    if (!cfg.pinned || !allowed.has(def.type) || widgets.some(w => w.type === def.type)) {
      return;
    }
    const template: IWidgetInstance | undefined = defaultLayout?.widgets.find(w => w.type === def.type);
    const instance: IWidgetInstance = apply(template ? { ...template, id: uid(def.type) } : { id: uid(def.type), type: def.type, size: def.defaultSize, settings: {} }, def);
    const index: number = defaultLayout ? defaultLayout.widgets.findIndex(w => w.type === def.type) : -1;
    widgets.splice(index >= 0 ? Math.min(index, widgets.length) : 0, 0, instance);
  });
  return { version: LAYOUT_VERSION, widgets, updated: state.updated };
}

// -----------------------------------------------------------------------------
// Provider
// -----------------------------------------------------------------------------

export const DashboardProvider: React.FC<IDashboardProviderProps> = (providerProps) => {
  const { context, graph, sp, theme, isDarkTheme, props, isPageEditMode, instanceId, onSaveAsDefault, children } = providerProps;
  const columns: number = Math.max(1, Math.min(4, props.columns || 3));
  const rulesKey: string = JSON.stringify([props.widgetConfig, props.enabledWidgets, props.defaultLayout, props.defaultLayouts]);

  const store: LayoutStore = React.useMemo(
    () =>
      new LayoutStore({
        mode: props.storageMode || 'local',
        listTitle: props.storageListTitle || 'ThePoint MySite Layouts',
        instanceKey: instanceId,
        userKey: context.pageContext.user.loginName || context.pageContext.user.email,
        graph,
        sp
      }),
    [props.storageMode, props.storageListTitle, instanceId, context, graph, sp]
  );
  const audienceService: AudienceService = React.useMemo(() => new AudienceService(graph), [graph]);

  const [audienceMap, setAudienceMap] = React.useState<Map<string, boolean> | undefined>(undefined);
  const [state, setState] = React.useState<IDashboardState>({ version: LAYOUT_VERSION, widgets: [] });
  const [ready, setReady] = React.useState<boolean>(false);
  const [loadSource, setLoadSource] = React.useState<'remote' | 'local' | 'none'>('none');
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [editMode, setEditModeState] = React.useState<boolean>(false);
  const [refreshToken, setRefreshToken] = React.useState<number>(0);
  const [unavailableTypes, setUnavailableTypes] = React.useState<Set<string>>(new Set());
  const stateRef = React.useRef<IDashboardState>(state);
  stateRef.current = state;

  // Every audience string the owner has used anywhere.
  const audienceStrings: string[] = React.useMemo(() => {
    const list: string[] = [];
    getWidgetDefinitions().forEach(def => {
      const a: string | undefined = widgetConfig(props, def.type).audiences;
      if (a && !AudienceService.isOpen(a)) {
        list.push(a);
      }
    });
    parseAudienceLayouts(props.defaultLayouts).forEach(e => {
      if (e.audiences && !AudienceService.isOpen(e.audiences)) {
        list.push(e.audiences);
      }
    });
    return list;
  }, [rulesKey]);

  const allowed: Set<string> = React.useMemo(() => computeAllowedTypes(props, audienceMap), [props, audienceMap]);
  const defaultLayout: IDashboardState = React.useMemo(() => resolveDefaultLayout(props, audienceMap), [props, audienceMap]);

  // Initial load: audiences and the saved layout in parallel.
  React.useEffect(() => {
    let active: boolean = true;
    setReady(false);
    Promise.all([audienceService.evaluate(audienceStrings), store.load()])
      .then(([map, result]) => {
        if (!active) {
          return;
        }
        const allowedNow: Set<string> = computeAllowedTypes(props, map);
        const defaults: IDashboardState = resolveDefaultLayout(props, map);
        setAudienceMap(map);
        setState(normalise(result.state || defaults, props, columns, allowedNow, defaults));
        setLoadSource(result.state ? result.source : 'none');
        setLoadError(result.error);
        setReady(true);
      })
      .catch((e: Error) => {
        if (active) {
          setLoadError(e.message);
          setReady(true);
        }
      });
    return () => {
      active = false;
    };
  }, [store, audienceService, rulesKey, columns]);

  // Permission probe: hide widgets whose Graph scope is not consented.
  React.useEffect(() => {
    if (!ready || !props.hideUnavailableWidgets) {
      setUnavailableTypes(new Set());
      return;
    }
    let active: boolean = true;
    const types: string[] = getWidgetDefinitions().filter(d => d.probe && allowed.has(d.type)).map(d => d.type);
    Promise.all(
      types.map(async type => {
        const def: IWidgetDefinition = getWidgetDefinition(type)!;
        const ok: boolean = await graph.cache.getOrFetch<boolean>(`probe:${type}`, async () => {
          try {
            await def.probe!(graph, sp);
            return true;
          } catch (e) {
            return !isForbidden(e);
          }
        });
        return { type, ok };
      })
    )
      .then(results => {
        if (active) {
          setUnavailableTypes(new Set(results.filter(r => !r.ok).map(r => r.type)));
        }
      })
      .catch(() => undefined);
    return () => {
      active = false;
    };
  }, [ready, props.hideUnavailableWidgets, graph, sp, allowed, refreshToken]);

  const commit = React.useCallback(
    (updater: (prev: IDashboardState) => IDashboardState): void => {
      setState(prev => {
        const next: IDashboardState = updater(prev);
        store.save(next);
        return next;
      });
    },
    [store]
  );

  const canCustomise: boolean = props.allowUserCustomisation !== false && !props.lockLayout;

  const setEditMode = React.useCallback((value: boolean): void => {
    setEditModeState(value && canCustomise);
    if (!value) {
      store.flush().catch(() => undefined);
    }
  }, [canCustomise, store]);

  const availableWidgets: IWidgetDefinition[] = React.useMemo(
    () => getWidgetDefinitions().filter(d => allowed.has(d.type) && !unavailableTypes.has(d.type)),
    [allowed, unavailableTypes]
  );

  const isPinned = React.useCallback((type: string): boolean => !!widgetConfig(props, type).pinned, [props]);
  const lockedKeys = React.useCallback((type: string): string[] => lockedSettingKeys(props, type), [props]);

  const addWidget = React.useCallback(
    (type: string, afterId?: string): void => {
      const def: IWidgetDefinition | undefined = getWidgetDefinition(type);
      if (!def || !allowed.has(type)) {
        return;
      }
      commit(prev => {
        if (def.allowMultiple === false && prev.widgets.some(w => w.type === type)) {
          return prev;
        }
        const instance: IWidgetInstance = {
          id: uid(type),
          type,
          size: Math.min(def.defaultSize, columns) as WidgetSize,
          settings: ownerDefaults(props, def)
        };
        const widgets: IWidgetInstance[] = [...prev.widgets];
        const index: number = afterId ? widgets.findIndex(w => w.id === afterId) : -1;
        if (index >= 0) {
          widgets.splice(index + 1, 0, instance);
        } else {
          widgets.push(instance);
        }
        return { ...prev, widgets };
      });
    },
    [commit, columns, allowed, props]
  );

  const removeWidget = React.useCallback(
    (id: string): void =>
      commit(prev => {
        const target: IWidgetInstance | undefined = prev.widgets.find(w => w.id === id);
        if (!target || isPinned(target.type)) {
          return prev;
        }
        return { ...prev, widgets: prev.widgets.filter(w => w.id !== id) };
      }),
    [commit, isPinned]
  );

  const moveWidget = React.useCallback(
    (id: string, toIndex: number): void =>
      commit(prev => {
        const from: number = prev.widgets.findIndex(w => w.id === id);
        if (from < 0) {
          return prev;
        }
        const widgets: IWidgetInstance[] = [...prev.widgets];
        const [item] = widgets.splice(from, 1);
        const target: number = Math.max(0, Math.min(widgets.length, toIndex > from ? toIndex - 1 : toIndex));
        widgets.splice(target, 0, item);
        return { ...prev, widgets };
      }),
    [commit]
  );

  const updateWidget = React.useCallback(
    (id: string, patch: Partial<IWidgetInstance>): void =>
      commit(prev => ({ ...prev, widgets: prev.widgets.map(w => (w.id === id ? { ...w, ...patch, id: w.id, type: w.type } : w)) })),
    [commit]
  );

  const updateWidgetSettings = React.useCallback(
    (id: string, patch: Record<string, unknown>): void =>
      commit(prev => ({
        ...prev,
        widgets: prev.widgets.map(w => {
          if (w.id !== id) {
            return w;
          }
          const clean: Record<string, unknown> = { ...patch };
          lockedSettingKeys(props, w.type).forEach(k => delete clean[k]);
          return { ...w, settings: { ...w.settings, ...clean } };
        })
      })),
    [commit, props]
  );

  const resetLayout = React.useCallback(async (): Promise<void> => {
    await store.reset();
    setState(normalise(defaultLayout, props, columns, allowed, defaultLayout));
    setLoadSource('none');
  }, [store, props, columns, allowed, defaultLayout]);

  const exportLayout = React.useCallback((): string => JSON.stringify({ version: LAYOUT_VERSION, widgets: stateRef.current.widgets }, undefined, 2), []);

  const importLayout = React.useCallback(
    (json: string): boolean => {
      const parsed: IDashboardState | undefined = parseLayout(json);
      if (!parsed) {
        return false;
      }
      commit(() => normalise(parsed, props, columns, allowed, defaultLayout));
      return true;
    },
    [commit, props, columns, allowed, defaultLayout]
  );

  const saveAsDefault = React.useMemo(
    () => (onSaveAsDefault && isPageEditMode ? (audience: string): void => onSaveAsDefault(exportLayout(), audience) : undefined),
    [onSaveAsDefault, isPageEditMode, exportLayout]
  );

  const refreshAll = React.useCallback((): void => {
    graph.cache.clear();
    sp.cache.clear();
    setRefreshToken(t => t + 1);
  }, [graph, sp]);

  const value: IDashboardContextValue = {
    context,
    graph,
    sp,
    theme,
    isDarkTheme,
    props,
    state,
    ready,
    loadSource,
    loadError,
    editMode,
    canCustomise,
    isPageEditMode,
    availableWidgets,
    unavailableTypes,
    isPinned,
    lockedKeys,
    setEditMode,
    addWidget,
    removeWidget,
    moveWidget,
    updateWidget,
    updateWidgetSettings,
    resetLayout,
    exportLayout,
    importLayout,
    saveAsDefault,
    refreshAll,
    refreshToken
  };

  return <DashboardContext.Provider value={value}>{children}</DashboardContext.Provider>;
};
