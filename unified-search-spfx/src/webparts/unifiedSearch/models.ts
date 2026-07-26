export type VerticalKey = 'all' | 'files' | 'mail' | 'calendar' | 'people' | 'sites' | 'policies';

export type FilterPanelPosition = 'left' | 'right' | 'hidden';

export type SearchScope = 'tenant' | 'hub' | 'site';

export type SortMode = 'relevance' | 'newest' | 'oldest';

export interface IUnifiedSearchConfig {
  filterPanelPosition: FilterPanelPosition;
  pageSize: number;
  showFiles: boolean;
  showMail: boolean;
  showCalendar: boolean;
  showPeople: boolean;
  showSites: boolean;
  showPolicies: boolean;
  enablePreviews: boolean;
  enableSuggestions: boolean;
  enableSpelling: boolean;
  openInNewTab: boolean;
  searchScope: SearchScope;
  policiesQuery: string;
}

export interface IFileResult {
  key: string;
  name: string;
  fileType: string;
  webUrl: string;
  driveId?: string;
  itemId?: string;
  siteName: string;
  modified?: string;
  modifiedBy?: string;
  author?: string;
  size?: number;
  /** Graph hit summary; contains <c0>…</c0> hit-highlight markers and <ddd/> ellipses. */
  summary?: string;
  isOneDrive: boolean;
}

export interface IMailResult {
  key: string;
  subject: string;
  from: string;
  fromEmail?: string;
  preview?: string;
  received?: string;
  hasAttachments: boolean;
  webLink?: string;
}

export interface IEventResult {
  key: string;
  subject: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: string;
  isOnline: boolean;
  webLink?: string;
}

export type PresenceAvailability =
  | 'Available'
  | 'Busy'
  | 'Away'
  | 'BeRightBack'
  | 'DoNotDisturb'
  | 'Offline'
  | 'PresenceUnknown';

export interface IPersonResult {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhone?: string;
  mobilePhone?: string;
  presence?: PresenceAvailability;
  photoUrl?: string;
}

export interface IPersonOrg {
  manager?: IPersonResult;
  directReports: IPersonResult[];
}

export interface ISiteResult {
  key: string;
  name: string;
  webUrl: string;
  description?: string;
}

export interface IRefinerBucket {
  key: string;
  label: string;
  count: number;
  /** Opaque token returned by Graph, echoed back verbatim in aggregationFilters. */
  filterToken: string;
}

export interface IRefiner {
  field: string;
  label: string;
  buckets: IRefinerBucket[];
}

export interface ISearchPage<T> {
  items: T[];
  total: number;
  moreAvailable: boolean;
  refiners?: IRefiner[];
  /** Spelling suggestion from Graph query alteration, when available. */
  spellingSuggestion?: string;
}

export interface ISuggestions {
  files: IFileResult[];
  people: IPersonResult[];
}

export interface IZeroState {
  recentSearches: string[];
  recentFiles: IFileResult[];
  upcomingEvents: IEventResult[];
}

export interface IVerticalDefinition {
  key: VerticalKey;
  label: string;
  icon: string;
}

export const VERTICALS: IVerticalDefinition[] = [
  { key: 'all', label: 'All', icon: 'Search' },
  { key: 'files', label: 'Files', icon: 'Documentation' },
  { key: 'mail', label: 'Mail', icon: 'Mail' },
  { key: 'calendar', label: 'Calendar', icon: 'Calendar' },
  { key: 'people', label: 'People', icon: 'People' },
  { key: 'sites', label: 'Sites', icon: 'Globe' },
  { key: 'policies', label: 'Policies', icon: 'Shield' }
];

export function isVerticalEnabled(config: IUnifiedSearchConfig, key: VerticalKey): boolean {
  switch (key) {
    case 'all': return true;
    case 'files': return config.showFiles;
    case 'mail': return config.showMail;
    case 'calendar': return config.showCalendar;
    case 'people': return config.showPeople;
    case 'sites': return config.showSites;
    case 'policies': return config.showPolicies;
    default: return false;
  }
}
