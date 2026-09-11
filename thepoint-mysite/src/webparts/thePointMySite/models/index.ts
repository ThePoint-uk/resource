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
  enabledWidgets: Record<string, boolean>;
  defaultLayout: string;
  maxItems: number;
  openLinksInNewTab: boolean;
  showPermissionHints: boolean;
}
