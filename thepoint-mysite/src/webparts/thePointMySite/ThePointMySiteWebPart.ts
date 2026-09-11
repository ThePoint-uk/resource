import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ThePointMySiteWebPartStrings';
import ThePointMySite from './components/ThePointMySite';
import type { IThePointMySiteProps } from './components/IThePointMySiteProps';
import type { IThePointMySiteWebPartProps } from './models';
import { GraphService } from './services/GraphService';
import { SharePointService } from './services/SharePointService';
import { CacheService } from './services/CacheService';
import { getWidgetDefinition, getWidgetDefinitions } from './widgets/registry';
import { isEnabled, parseAudienceLayouts, schemaDefaults } from './context/DashboardContext';
import type { IAudienceLayout, IWidgetDefinition } from './models';

export type { IThePointMySiteWebPartProps };

/**
 * ThePoint MySite — a personalised, widget-based dashboard for SharePoint,
 * Teams and Microsoft 365. Users pick, arrange and configure widgets;
 * page authors control the defaults, look and feel, and which widgets are available.
 */
export default class ThePointMySiteWebPart extends BaseClientSideWebPart<IThePointMySiteWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _theme: IReadonlyTheme | undefined;
  private _cache!: CacheService;
  private _graph!: GraphService;
  private _sp!: SharePointService;
  private _ruleWidget: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._applyDefaults();
    this._cache = new CacheService(this.context.pageContext.user.loginName || 'anon', this.properties.cacheMinutes);
    this._graph = new GraphService(this.context, this._cache);
    this._sp = new SharePointService(this.context, this._cache);
  }

  public render(): void {
    this._cache.ttlMinutes = this.properties.cacheMinutes;
    const element: React.ReactElement<IThePointMySiteProps> = React.createElement(ThePointMySite, {
      context: this.context,
      graph: this._graph,
      sp: this._sp,
      theme: this._theme,
      isDarkTheme: this._isDarkTheme,
      properties: { ...this.properties },
      isPageEditMode: this.displayMode === DisplayMode.Edit,
      instanceId: this.instanceId,
      onSaveAsDefault: (json: string, audience: string): void => {
        if (!audience) {
          this.properties.defaultLayout = json;
        } else {
          const layouts: IAudienceLayout[] = parseAudienceLayouts(this.properties.defaultLayouts);
          const existing: IAudienceLayout | undefined = layouts.find(l => l.audiences.trim().toLowerCase() === audience.toLowerCase());
          if (existing) {
            existing.layout = JSON.parse(json);
          } else {
            layouts.push({ audiences: audience, layout: JSON.parse(json) });
          }
          this.properties.defaultLayouts = JSON.stringify(layouts, undefined, 2);
        }
        this.context.propertyPane.refresh();
        this.render();
      }
    });
    ReactDom.render(element, this.domElement);
  }

  private _applyDefaults(): void {
    const p: IThePointMySiteWebPartProps = this.properties;
    p.title = p.title ?? 'My workspace';
    p.showTitle = p.showTitle ?? false;
    p.showGreeting = p.showGreeting ?? true;
    p.greetingStyle = p.greetingStyle ?? 'timeOfDay';
    p.showDate = p.showDate ?? true;
    p.showSearch = p.showSearch ?? true;
    p.columns = p.columns ?? 3;
    p.density = p.density ?? 'comfortable';
    p.cardStyle = p.cardStyle ?? 'shadow';
    p.cornerRadius = p.cornerRadius ?? 8;
    p.accentColor = p.accentColor ?? '';
    p.headerStyle = p.headerStyle ?? 'plain';
    p.allowUserCustomisation = p.allowUserCustomisation ?? true;
    p.lockLayout = p.lockLayout ?? false;
    p.storageMode = p.storageMode ?? 'local';
    p.storageListTitle = p.storageListTitle ?? 'ThePoint MySite Layouts';
    p.cacheMinutes = p.cacheMinutes ?? 10;
    p.widgetConfig = p.widgetConfig ?? {};
    // Migrate the v1.0 enable map into widgetConfig.
    if (p.enabledWidgets) {
      Object.keys(p.enabledWidgets).forEach(type => {
        if (p.enabledWidgets![type] === false && p.widgetConfig[type]?.enabled === undefined) {
          p.widgetConfig[type] = { ...(p.widgetConfig[type] || {}), enabled: false };
        }
      });
      p.enabledWidgets = undefined;
    }
    p.defaultLayout = p.defaultLayout ?? '';
    p.defaultLayouts = p.defaultLayouts ?? '';
    p.hideUnavailableWidgets = p.hideUnavailableWidgets ?? true;
    p.maxItems = p.maxItems ?? 10;
    p.openLinksInNewTab = p.openLinksInNewTab ?? true;
    p.showPermissionHints = p.showPermissionHints ?? true;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._theme = currentTheme;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'ruleWidget') {
      // The rule selector is UI state, not a persisted property.
      this._ruleWidget = String(newValue);
      delete (this.properties as unknown as Record<string, unknown>).ruleWidget;
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const definitions: IWidgetDefinition[] = getWidgetDefinitions();
    const widgetToggles: IPropertyPaneField<unknown>[] = definitions.map(def =>
      PropertyPaneCheckbox(`widgetConfig.${def.type}.enabled`, {
        text: def.title,
        checked: isEnabled(this.properties, def.type)
      })
    );

    const ruleType: string = this._ruleWidget && getWidgetDefinition(this._ruleWidget) ? this._ruleWidget : definitions[0].type;
    const ruleDef: IWidgetDefinition = getWidgetDefinition(ruleType)!;
    const ruleFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('ruleWidget', {
        label: strings.RuleWidgetLabel,
        options: definitions.map(def => ({ key: def.type, text: def.title })),
        selectedKey: ruleType
      }),
      PropertyPaneToggle(`widgetConfig.${ruleType}.pinned`, { label: strings.PinnedLabel, onText: strings.PinnedOn, offText: strings.PinnedOff }),
      PropertyPaneTextField(`widgetConfig.${ruleType}.audiences`, {
        label: strings.AudiencesLabel,
        description: strings.AudiencesDescription,
        placeholder: strings.AudiencesPlaceholder
      })
    ];
    const lockable = (ruleDef.settingsSchema || []).filter(f => f.type !== 'links');
    if (lockable.length > 0) {
      ruleFields.push(PropertyPaneLabel('lockedHelp', { text: strings.LockedSettingsHelp }));
      lockable.forEach(f => {
        ruleFields.push(
          PropertyPaneCheckbox(`widgetConfig.${ruleType}.locked.${f.key}`, {
            text: f.label,
            checked: !!this.properties.widgetConfig?.[ruleType]?.locked?.[f.key]
          })
        );
      });
      ruleFields.push(
        PropertyPaneTextField(`widgetConfig.${ruleType}.settings`, {
          label: strings.OwnerSettingsLabel,
          description: strings.OwnerSettingsDescription,
          multiline: true,
          rows: 6,
          resizable: true,
          onGetErrorMessage: (value: string): string => {
            if (!value || !value.trim()) {
              return '';
            }
            try {
              const parsed: unknown = JSON.parse(value);
              return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? '' : strings.OwnerSettingsError;
            } catch {
              return strings.OwnerSettingsError;
            }
          }
        }),
        PropertyPaneButton('fillDefaults', {
          text: strings.FillDefaultsButton,
          buttonType: PropertyPaneButtonType.Normal,
          icon: 'Refresh',
          onClick: (): void => {
            this.properties.widgetConfig[ruleType] = {
              ...(this.properties.widgetConfig[ruleType] || {}),
              settings: JSON.stringify(schemaDefaults(ruleDef), undefined, 2)
            };
            this.context.propertyPane.refresh();
            this.render();
          }
        })
      );
    }

    return {
      pages: [
        {
          header: { description: strings.PageGeneralDescription },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.GroupHeader,
              groupFields: [
                PropertyPaneToggle('showGreeting', { label: strings.ShowGreetingLabel }),
                PropertyPaneDropdown('greetingStyle', {
                  label: strings.GreetingStyleLabel,
                  options: [
                    { key: 'timeOfDay', text: strings.GreetingTimeOfDay },
                    { key: 'welcome', text: strings.GreetingWelcome },
                    { key: 'none', text: strings.GreetingNone }
                  ],
                  disabled: !this.properties.showGreeting
                }),
                PropertyPaneToggle('showTitle', { label: strings.ShowTitleLabel }),
                PropertyPaneTextField('title', { label: strings.TitleLabel, disabled: !this.properties.showTitle }),
                PropertyPaneToggle('showDate', { label: strings.ShowDateLabel }),
                PropertyPaneToggle('showSearch', { label: strings.ShowSearchLabel }),
                PropertyPaneChoiceGroup('headerStyle', {
                  label: strings.HeaderStyleLabel,
                  options: [
                    { key: 'plain', text: strings.HeaderPlain, iconProps: { officeFabricIconFontName: 'TextField' } },
                    { key: 'accent', text: strings.HeaderAccent, iconProps: { officeFabricIconFontName: 'Color' } },
                    { key: 'gradient', text: strings.HeaderGradient, iconProps: { officeFabricIconFontName: 'BackgroundColor' } }
                  ]
                })
              ]
            },
            {
              groupName: strings.GroupLayout,
              groupFields: [
                PropertyPaneSlider('columns', { label: strings.ColumnsLabel, min: 1, max: 4, step: 1, showValue: true }),
                PropertyPaneChoiceGroup('density', {
                  label: strings.DensityLabel,
                  options: [
                    { key: 'comfortable', text: strings.DensityComfortable },
                    { key: 'compact', text: strings.DensityCompact }
                  ]
                }),
                PropertyPaneChoiceGroup('cardStyle', {
                  label: strings.CardStyleLabel,
                  options: [
                    { key: 'shadow', text: strings.CardShadow },
                    { key: 'outline', text: strings.CardOutline },
                    { key: 'flat', text: strings.CardFlat }
                  ]
                }),
                PropertyPaneSlider('cornerRadius', { label: strings.CornerRadiusLabel, min: 0, max: 20, step: 1, showValue: true }),
                PropertyPaneTextField('accentColor', {
                  label: strings.AccentColorLabel,
                  description: strings.AccentColorDescription,
                  placeholder: '#0078d4',
                  onGetErrorMessage: (value: string): string => (!value || /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(value) ? '' : strings.AccentColorError)
                }),
                PropertyPaneToggle('openLinksInNewTab', { label: strings.OpenLinksInNewTabLabel })
              ]
            }
          ]
        },
        {
          header: { description: strings.PageWidgetsDescription },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.GroupAvailableWidgets,
              groupFields: [
                PropertyPaneLabel('widgetsHelp', { text: strings.AvailableWidgetsHelp }),
                ...widgetToggles,
                PropertyPaneToggle('hideUnavailableWidgets', { label: strings.HideUnavailableLabel }),
                PropertyPaneToggle('showPermissionHints', { label: strings.ShowPermissionHintsLabel })
              ]
            },
            {
              groupName: strings.GroupWidgetRules,
              groupFields: [PropertyPaneLabel('rulesHelp', { text: strings.WidgetRulesHelp }), ...ruleFields]
            },
            {
              groupName: strings.GroupDefaultLayout,
              groupFields: [
                PropertyPaneLabel('defaultLayoutHelp', { text: strings.DefaultLayoutHelp }),
                PropertyPaneTextField('defaultLayout', {
                  label: strings.DefaultLayoutLabel,
                  multiline: true,
                  rows: 10,
                  resizable: true,
                  onGetErrorMessage: (value: string): string => {
                    if (!value || !value.trim()) {
                      return '';
                    }
                    try {
                      const parsed: { widgets?: unknown } = JSON.parse(value);
                      return Array.isArray(parsed.widgets) ? '' : strings.DefaultLayoutError;
                    } catch {
                      return strings.DefaultLayoutError;
                    }
                  }
                }),
                PropertyPaneButton('clearDefaultLayout', {
                  text: strings.ClearDefaultLayoutButton,
                  buttonType: PropertyPaneButtonType.Normal,
                  icon: 'Clear',
                  onClick: (): void => {
                    this.properties.defaultLayout = '';
                    this.context.propertyPane.refresh();
                    this.render();
                  }
                }),
                PropertyPaneLabel('audienceLayoutsHelp', { text: strings.AudienceLayoutsHelp }),
                PropertyPaneTextField('defaultLayouts', {
                  label: strings.AudienceLayoutsLabel,
                  multiline: true,
                  rows: 8,
                  resizable: true,
                  onGetErrorMessage: (value: string): string => {
                    if (!value || !value.trim()) {
                      return '';
                    }
                    try {
                      const parsed: unknown = JSON.parse(value);
                      return Array.isArray(parsed) ? '' : strings.AudienceLayoutsError;
                    } catch {
                      return strings.AudienceLayoutsError;
                    }
                  }
                })
              ]
            }
          ]
        },
        {
          header: { description: strings.PagePersonalisationDescription },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.GroupPersonalisation,
              groupFields: [
                PropertyPaneToggle('allowUserCustomisation', { label: strings.AllowCustomisationLabel }),
                PropertyPaneToggle('lockLayout', { label: strings.LockLayoutLabel, disabled: !this.properties.allowUserCustomisation }),
                PropertyPaneDropdown('storageMode', {
                  label: strings.StorageModeLabel,
                  options: [
                    { key: 'local', text: strings.StorageLocal },
                    { key: 'onedrive', text: strings.StorageOneDrive },
                    { key: 'list', text: strings.StorageList }
                  ]
                }),
                PropertyPaneLabel('storageHelp', {
                  text:
                    this.properties.storageMode === 'onedrive'
                      ? strings.StorageOneDriveHelp
                      : this.properties.storageMode === 'list'
                        ? strings.StorageListHelp
                        : strings.StorageLocalHelp
                }),
                PropertyPaneTextField('storageListTitle', { label: strings.StorageListTitleLabel, disabled: this.properties.storageMode !== 'list' })
              ]
            },
            {
              groupName: strings.GroupPerformance,
              groupFields: [
                PropertyPaneSlider('cacheMinutes', { label: strings.CacheMinutesLabel, min: 0, max: 120, step: 5, showValue: true }),
                PropertyPaneLabel('cacheHelp', { text: strings.CacheMinutesHelp })
              ]
            },
            {
              groupName: strings.GroupAbout,
              groupFields: [PropertyPaneLabel('about', { text: strings.AboutText })]
            }
          ]
        }
      ]
    };
  }
}
