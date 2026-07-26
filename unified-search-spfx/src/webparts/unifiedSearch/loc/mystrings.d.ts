declare interface IUnifiedSearchWebPartStrings {
  PropertyPaneDescription: string;
  LayoutGroupName: string;
  VerticalsGroupName: string;
  BehaviourGroupName: string;
  CopilotGroupName: string;
  FilterPanelPositionLabel: string;
  PageSizeLabel: string;
  ShowFilesLabel: string;
  ShowMailLabel: string;
  ShowCalendarLabel: string;
  ShowPeopleLabel: string;
  ShowSitesLabel: string;
  ShowPoliciesLabel: string;
  EnablePreviewsLabel: string;
  EnableSuggestionsLabel: string;
  EnableSpellingLabel: string;
  OpenInNewTabLabel: string;
  SearchScopeLabel: string;
  PoliciesQueryLabel: string;
  PoliciesQueryDescription: string;
  CopilotRoadmapText: string;
}

declare module 'UnifiedSearchWebPartStrings' {
  const strings: IUnifiedSearchWebPartStrings;
  export = strings;
}
