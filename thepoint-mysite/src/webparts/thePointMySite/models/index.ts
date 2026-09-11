import type * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { GraphService } from '../services/GraphService';
import type { SharePointService } from '../services/SharePointService';

/** Column span options for a widget inside the dashboard grid. */
export type WidgetSize = 1 | 2 | 3 | 4;

/** Widget categories used by the catalogue. */
export type WidgetCategory = 'me' | 'people' | 'content' | 'productivity' | 'personal';

/** Persistence strategies for the per-user layout. */
export type StorageMode = 'local' | 'onedrive' | 'list';

/** Card visual styles configurable by the page author. */
export type CardStyle = 'shadow' | 'outline' | 'flat';

/** Card density options. */
export type Density = 'comfortable' | 'compact';

/**
 * A single widget placed on the user's dashboard.
 */
export interface IWidgetInstance {
  /** Unique instance id (a widget type can be placed more than once). */
  id: string;
  /** Registry key of the widget definition. */
  type: string;
  /** Column span. */
  size: WidgetSize;
  /** Collapsed state. */
  collapsed?: boolean;
  /** Custom title, overrides the definition title when set. */
  title?: string;
  /** Per-instance settings, validated against the definition's settings schema. */
  settings: Record<string, unknown>;
}

/**
 * The full persisted dashboard state for a user.
 */
export interface IDashboardState {
  version: number;
  widgets: IWidgetInstance[];
  updated?: string;
}

/**
 * Setting field types supported by the generic widget settings form.
 */
export type WidgetSettingType = 'text' | 'number' | 'boolean' | 'choice' | 'multiline' | 'links';

export interface IWidgetSettingChoice {
  key: string;
  text: string;
}

export interface IWidgetSettingField {
  key: string;
  label: string;
  type: WidgetSettingType;
  description?: string;
  defaultValue?: unknown;
  min?: number;
  max?: number;
  step?: number;
  choices?: IWidgetSettingChoice[];
}

/**
 * Props passed to every widget component.
 */
export interface IWidgetProps {
  instance: IWidgetInstance;
  settings: Record<string, unknown>;
  context: WebPartContext;
  graph: GraphService;
  sp: SharePointService;
  theme: IReadonlyTheme | undefined;
  isDarkTheme: boolean;
  density: Density;
  editMode: boolean;
  /** Updates settings for this instance and persists them. */
  updateSettings: (patch: Record<string, unknown>) => void;
  /** Column span currently in use (useful for responsive rendering). */
  size: WidgetSize;
}

/**
 * Owner-side rules for one widget type, set in the property pane.
 */
export interface IWidgetConfig {
  /** Available in the catalogue and allowed in layouts. Defaults to true. */
  enabled?: boolean;
  /** Users can move and resize but not remove. */
  pinned?: boolean;
  /** Comma separated Entra / Microsoft 365 group names, ids or mail addresses; blank = everyone. Supports @everyone and @managers. */
  audiences?: string;
  /** Setting keys users may not change; the owner value (or schema default) always wins. */
  locked?: Record<string, boolean>;
  /** Owner default settings as JSON; applied to new instances and to locked keys. */
  settings?: string;
}

/**
 * A default layout that applies to a specific audience.
 */
export interface IAudienceLayout {
  audiences: string;
  layout: IDashboardState;
}

/**
 * Static definition of a widget type.
 */
export interface IWidgetDefinition {
  type: string;
  title: string;
  description: string;
  icon: string;
  category: WidgetCategory;
  defaultSize: WidgetSize;
  /** Allowed sizes; defaults to all. */
  sizes?: WidgetSize[];
  /** Whether more than one instance is allowed. */
  allowMultiple?: boolean;
  /** Graph / SharePoint permissions this widget relies on, shown in the catalogue. */
  permissions?: string[];
  settingsSchema?: IWidgetSettingField[];
  /**
   * Cheap call that fails with 403 when the widget's permission has not been consented.
   * Used to hide widgets the user cannot use.
   */
  probe?: (graph: GraphService, sp: SharePointService) => Promise<void>;
  component: React.ComponentType<IWidgetProps>;
}

/**
 * Quick link entry used by the Quick Links widget.
 */
export interface IQuickLink {
  id: string;
  title: string;
  url: string;
  icon?: string;
}

/**
 * Web part properties set by the page author via the property pane.
 */
export interface IThePointMySiteWebPartProps {
  title: string;
  showTitle: boolean;
  showGreeting: boolean;
  greetingStyle: 'timeOfDay' | 'welcome' | 'none';
  showDate: boolean;
  showSearch: boolean;
  columns: number;
  density: Density;
  cardStyle: CardStyle;
  cornerRadius: number;
  accentColor: string;
  headerStyle: 'plain' | 'accent' | 'gradient';
  allowUserCustomisation: boolean;
  lockLayout: boolean;
  storageMode: StorageMode;
  storageListTitle: string;
  cacheMinutes: number;
  /** Legacy (v1.0) enable map, migrated into widgetConfig on init. */
  enabledWidgets?: Record<string, boolean>;
  widgetConfig: Record<string, IWidgetConfig>;
  defaultLayout: string;
  /** JSON array of IAudienceLayout, evaluated in order; first match wins. */
  defaultLayouts: string;
  hideUnavailableWidgets: boolean;
  maxItems: number;
  openLinksInNewTab: boolean;
  showPermissionHints: boolean;
}
