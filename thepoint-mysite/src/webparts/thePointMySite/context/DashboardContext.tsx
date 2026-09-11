import * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { IDashboardState, IWidgetDefinition, IWidgetInstance, IThePointMySiteWebPartProps, WidgetSize } from '../models';
import { LayoutStore, LAYOUT_VERSION } from '../services/LayoutStore';
import type { GraphService } from '../services/GraphService';
import type { SharePointService } from '../services/SharePointService';
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
  availableWidgets: IWidgetDefinition[];
  setEditMode: (value: boolean) => void;
  addWidget: (type: string, afterId?: string) => void;
  removeWidget: (id: string) => void;
  moveWidget: (id: string, toIndex: number) => void;
  updateWidget: (id: string, patch: Partial<IWidgetInstance>) => void;
  updateWidgetSettings: (id: string, patch: Record<string, unknown>) => void;
  resetLayout: () => Promise<void>;
  exportLayout: () => string;
  importLayout: (json: string) => boolean;
  saveAsDefault: (() => void) | undefined;
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
  onSaveAsDefault?: (json: string) => void;
  children?: React.ReactNode;
}

/**
 * Builds the default layout from the web part's `defaultLayout` JSON (if valid),
 * otherwise a sensible starter set of widgets.
 */
export function buildDefaultState(props: IThePointMySiteWebPartProps, columns: number): IDashboardState {
  const parsed: IDashboardState | undefined = parseLayout(props.defaultLayout);
  if (parsed) {
    return normalise(parsed, props, columns);
  }
  const starter: { type: string; size: WidgetSize }[] = [
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
  return normalise(
    {
      version: LAYOUT_VERSION,
      widgets: starter.map(s => ({ id: uid(s.type), type: s.type, size: s.size, settings: {} }))
    },
    props,
    columns
  );
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

/** Drops unknown/disabled widgets, clamps sizes and fills missing settings from defaults. */
export function normalise(state: IDashboardState, props: IThePointMySiteWebPartProps, columns: number): IDashboardState {
  const seen: Set<string> = new Set();
  const widgets: IWidgetInstance[] = [];
  for (const w of state.widgets) {
    const def: IWidgetDefinition | undefined = getWidgetDefinition(w.type);
    if (!def || props.enabledWidgets?.[w.type] === false) {
      continue;
    }
    let id: string = w.id || uid(w.type);
    if (seen.has(id)) {
      id = uid(w.type);
    }
    seen.add(id);
    const settings: Record<string, unknown> = { ...defaultSettings(def), ...(w.settings || {}) };
    const size: WidgetSize = Math.max(1, Math.min(columns, Number(w.size) || def.defaultSize)) as WidgetSize;
    widgets.push({ id, type: w.type, size, collapsed: !!w.collapsed, title: w.title, settings });
  }
  return { version: LAYOUT_VERSION, widgets, updated: state.updated };
}

export function defaultSettings(def: IWidgetDefinition): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  (def.settingsSchema || []).forEach(field => {
    if (field.defaultValue !== undefined) {
      result[field.key] = field.defaultValue;
    }
  });
  return result;
}

export const DashboardProvider: React.FC<IDashboardProviderProps> = (providerProps) => {
  const { context, graph, sp, theme, isDarkTheme, props, isPageEditMode, instanceId, onSaveAsDefault, children } = providerProps;
  const columns: number = Math.max(1, Math.min(4, props.columns || 3));

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

  const [state, setState] = React.useState<IDashboardState>(() => buildDefaultState(props, columns));
  const [ready, setReady] = React.useState<boolean>(false);
  const [loadSource, setLoadSource] = React.useState<'remote' | 'local' | 'none'>('none');
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [editMode, setEditModeState] = React.useState<boolean>(false);
  const [refreshToken, setRefreshToken] = React.useState<number>(0);
  const stateRef = React.useRef<IDashboardState>(state);
  stateRef.current = state;

  // Initial load
  React.useEffect(() => {
    let active: boolean = true;
    setReady(false);
    store
      .load()
      .then(result => {
        if (!active) {
          return;
        }
        if (result.state) {
          setState(normalise(result.state, props, columns));
        } else {
          setState(buildDefaultState(props, columns));
        }
        setLoadSource(result.source);
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
    // Re-run when the store or the author-controlled defaults change.
  }, [store, props.defaultLayout, columns]);

  // Re-normalise when author toggles enabled widgets.
  React.useEffect(() => {
    if (ready) {
      setState(prev => normalise(prev, props, columns));
    }
  }, [JSON.stringify(props.enabledWidgets)]);

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
    () => getWidgetDefinitions().filter(d => props.enabledWidgets?.[d.type] !== false),
    [props.enabledWidgets]
  );

  const addWidget = React.useCallback(
    (type: string, afterId?: string): void => {
      const def: IWidgetDefinition | undefined = getWidgetDefinition(type);
      if (!def) {
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
          settings: defaultSettings(def)
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
    [commit, columns]
  );

  const removeWidget = React.useCallback(
    (id: string): void => commit(prev => ({ ...prev, widgets: prev.widgets.filter(w => w.id !== id) })),
    [commit]
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
        widgets: prev.widgets.map(w => (w.id === id ? { ...w, settings: { ...w.settings, ...patch } } : w))
      })),
    [commit]
  );

  const resetLayout = React.useCallback(async (): Promise<void> => {
    await store.reset();
    setState(buildDefaultState(props, columns));
    setLoadSource('none');
  }, [store, props, columns]);

  const exportLayout = React.useCallback((): string => JSON.stringify({ version: LAYOUT_VERSION, widgets: stateRef.current.widgets }, undefined, 2), []);

  const importLayout = React.useCallback(
    (json: string): boolean => {
      const parsed: IDashboardState | undefined = parseLayout(json);
      if (!parsed) {
        return false;
      }
      commit(() => normalise(parsed, props, columns));
      return true;
    },
    [commit, props, columns]
  );

  const saveAsDefault = React.useMemo(
    () => (onSaveAsDefault && isPageEditMode ? (): void => onSaveAsDefault(exportLayout()) : undefined),
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
