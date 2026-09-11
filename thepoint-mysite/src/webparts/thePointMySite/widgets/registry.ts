import type { IWidgetDefinition, WidgetCategory } from '../models';
import { myProfileDefinition } from './profile/MyProfileWidget';
import { myTeamDefinition } from './team/MyTeamWidget';
import { myOrgDefinition } from './org/MyOrgWidget';
import { peopleDefinition } from './people/PeopleWidget';
import { myAppsDefinition } from './apps/MyAppsWidget';
import { recentDocumentsDefinition } from './recentDocs/RecentDocumentsWidget';
import { sitesDefinition } from './sites/SitesWidget';
import { oneDriveDefinition } from './onedrive/OneDriveWidget';
import { calendarDefinition } from './calendar/CalendarWidget';
import { tasksDefinition } from './tasks/TasksWidget';
import { mailDefinition } from './mail/MailWidget';
import { teamsDefinition } from './teams/TeamsWidget';
import { newsDefinition } from './news/NewsWidget';
import { quickLinksDefinition } from './quickLinks/QuickLinksWidget';
import { notesDefinition } from './notes/NotesWidget';

/**
 * Registry of every widget the dashboard can render. Order controls catalogue order.
 * To add a widget: create a definition exporting `IWidgetDefinition` and push it here.
 */
const REGISTRY: IWidgetDefinition[] = [
  myProfileDefinition,
  myTeamDefinition,
  myOrgDefinition,
  peopleDefinition,
  teamsDefinition,
  myAppsDefinition,
  recentDocumentsDefinition,
  sitesDefinition,
  oneDriveDefinition,
  newsDefinition,
  calendarDefinition,
  tasksDefinition,
  mailDefinition,
  quickLinksDefinition,
  notesDefinition
];

const BY_TYPE: Map<string, IWidgetDefinition> = new Map(REGISTRY.map(d => [d.type, d]));

export function getWidgetDefinitions(): IWidgetDefinition[] {
  return REGISTRY;
}

export function getWidgetDefinition(type: string): IWidgetDefinition | undefined {
  return BY_TYPE.get(type);
}

export const CATEGORY_LABELS: Record<WidgetCategory, string> = {
  me: 'Me',
  people: 'People',
  content: 'Content',
  productivity: 'Productivity',
  personal: 'Personal'
};

export const CATEGORY_ORDER: WidgetCategory[] = ['me', 'people', 'content', 'productivity', 'personal'];
