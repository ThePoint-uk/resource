define([], function () {
  return {
    PropertyPaneDescription:
      'Configure the unified search experience. All queries run as the signed-in user via Microsoft Graph delegated search, so results always respect existing permissions.',
    LayoutGroupName: 'Layout',
    VerticalsGroupName: 'Result types',
    BehaviourGroupName: 'Search behaviour',
    CopilotGroupName: 'Copilot & AI (roadmap)',
    FilterPanelPositionLabel: 'Filter panel position',
    PageSizeLabel: 'Results per page',
    ShowFilesLabel: 'Files (SharePoint & OneDrive)',
    ShowMailLabel: 'Mail',
    ShowCalendarLabel: 'Calendar',
    ShowPeopleLabel: 'People',
    ShowSitesLabel: 'Sites',
    ShowPoliciesLabel: 'Policies',
    EnablePreviewsLabel: 'File previews',
    EnableSuggestionsLabel: 'Query suggestions',
    EnableSpellingLabel: 'Spelling suggestions ("Did you mean…")',
    OpenInNewTabLabel: 'Open results in a new tab',
    SearchScopeLabel: 'Search scope',
    PoliciesQueryLabel: 'Policies query template (KQL)',
    PoliciesQueryDescription:
      'KQL appended to the user query on the Policies tab, e.g. a content type or a path restriction to your policy library.',
    PromotedGroupName: 'Promoted results',
    EnablePromotedLabel: 'Show promoted results',
    PromotedListTitleLabel: 'Promoted results list',
    PromotedListTitleDescription:
      'Display name of the SharePoint list that holds promoted results. Content owners edit the list; no redeployment needed.',
    PromotedWebUrlLabel: 'List site URL (optional)',
    PromotedWebUrlDescription:
      'Absolute URL of the site that hosts the list. Leave empty to use the current site.',
    CopilotRoadmapText:
      'Planned: a grounded Copilot answer card with citations, natural-language filters and chat-with-results via the Microsoft Graph / Copilot Retrieval API. These stay inside the tenant security boundary and respect the same permissions as search. Not included in v1.'
  };
});
