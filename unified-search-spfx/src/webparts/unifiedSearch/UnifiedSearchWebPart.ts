import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'UnifiedSearchWebPartStrings';

import { UnifiedSearch } from './components/UnifiedSearch';
import { IUnifiedSearchConfig } from './models';
import { GraphSearchService } from './services/GraphSearchService';
import { PromotedResultsService } from './services/PromotedResultsService';
import { PeopleService } from './services/PeopleService';
import { FileService } from './services/FileService';

export interface IUnifiedSearchWebPartProps extends IUnifiedSearchConfig { }

export default class UnifiedSearchWebPart extends BaseClientSideWebPart<IUnifiedSearchWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _theme: IReadonlyTheme | undefined;
  private _searchService: GraphSearchService;
  private _peopleService: PeopleService;
  private _fileService: FileService;
  private _promotedService: PromotedResultsService;

  protected async onInit(): Promise<void> {
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._onThemeChanged);

    const factory = this.context.msGraphClientFactory;
    const site = this.context.pageContext.site;
    const legacy = this.context.pageContext.legacyPageContext as { hubSiteId?: string } | undefined;
    this._searchService = new GraphSearchService(
      factory,
      site.absoluteUrl,
      legacy && legacy.hubSiteId ? legacy.hubSiteId : undefined
    );
    this._peopleService = new PeopleService(factory);
    this._fileService = new FileService(factory);
    this._promotedService = new PromotedResultsService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
  }

  public render(): void {
    const element = React.createElement(UnifiedSearch, {
      config: { ...this.properties },
      theme: this._theme,
      searchService: this._searchService,
      peopleService: this._peopleService,
      fileService: this._fileService,
      promotedService: this._promotedService,
      userDisplayName: this.context.pageContext.user.displayName,
      isEditMode: this.displayMode === 2
    });
    ReactDom.render(element, this.domElement);
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneDropdown('filterPanelPosition', {
                  label: strings.FilterPanelPositionLabel,
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'right', text: 'Right' },
                    { key: 'hidden', text: 'Hidden' }
                  ]
                }),
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeLabel,
                  min: 5,
                  max: 30,
                  step: 5
                })
              ]
            },
            {
              groupName: strings.VerticalsGroupName,
              groupFields: [
                PropertyPaneToggle('showFiles', { label: strings.ShowFilesLabel }),
                PropertyPaneToggle('showMail', { label: strings.ShowMailLabel }),
                PropertyPaneToggle('showCalendar', { label: strings.ShowCalendarLabel }),
                PropertyPaneToggle('showPeople', { label: strings.ShowPeopleLabel }),
                PropertyPaneToggle('showSites', { label: strings.ShowSitesLabel }),
                PropertyPaneToggle('showPolicies', { label: strings.ShowPoliciesLabel }),
                PropertyPaneTextField('policiesQuery', {
                  label: strings.PoliciesQueryLabel,
                  description: strings.PoliciesQueryDescription,
                  multiline: true,
                  rows: 2
                })
              ]
            },
            {
              groupName: strings.PromotedGroupName,
              groupFields: [
                PropertyPaneToggle('enablePromoted', { label: strings.EnablePromotedLabel }),
                PropertyPaneTextField('promotedListTitle', {
                  label: strings.PromotedListTitleLabel,
                  description: strings.PromotedListTitleDescription
                }),
                PropertyPaneTextField('promotedWebUrl', {
                  label: strings.PromotedWebUrlLabel,
                  description: strings.PromotedWebUrlDescription
                })
              ]
            },
            {
              groupName: strings.BehaviourGroupName,
              groupFields: [
                PropertyPaneDropdown('searchScope', {
                  label: strings.SearchScopeLabel,
                  options: [
                    { key: 'tenant', text: 'Entire organisation' },
                    { key: 'hub', text: 'Current hub' },
                    { key: 'site', text: 'Current site' }
                  ]
                }),
                PropertyPaneToggle('enablePreviews', { label: strings.EnablePreviewsLabel }),
                PropertyPaneToggle('enableSuggestions', { label: strings.EnableSuggestionsLabel }),
                PropertyPaneToggle('enableSpelling', { label: strings.EnableSpellingLabel }),
                PropertyPaneToggle('openInNewTab', { label: strings.OpenInNewTabLabel })
              ]
            },
            {
              groupName: strings.CopilotGroupName,
              groupFields: [
                PropertyPaneLabel('copilotRoadmap', { text: strings.CopilotRoadmapText })
              ]
            }
          ]
        }
      ]
    };
  }
}
