import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'StoragePulseWebPartStrings';
import { StoragePulse } from './components/StoragePulse';
import { IStoragePulseProps, IStoragePulseTheme, ScanMode, ScanPermission } from './components/IStoragePulseProps';
import { format, thresholdLabel } from './components/text';
import { ScanScope } from './models/IScanResult';
import { DEFAULT_THRESHOLD_MONTHS, THRESHOLD_OPTIONS } from './services/activity';
import { ScanSpeed } from './services/RequestGovernor';

export interface IStoragePulseWebPartProps {
  title: string;
  scope: ScanScope;
  thresholdMonths: number;
  scanPermission: ScanPermission;
  scanMode: ScanMode;
  scanSpeed: ScanSpeed;
  includeHidden: boolean;
  excludeSystemLibraries: boolean;
  /** One library title or URL name per line. */
  excludedLibraries: string;
  /** 0 turns the stale-results warning off. */
  staleAfterDays: number;
}

function luminance(color: string | undefined): number | undefined {
  const match = color ? /^#([0-9a-f]{6})$/i.exec(color.trim()) : null;
  if (!match) {
    return undefined;
  }
  const value = parseInt(match[1], 16);
  const channel = (c: number): number => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel((value >> 16) & 255) + 0.7152 * channel((value >> 8) & 255) + 0.0722 * channel(value & 255);
}

export default class StoragePulseWebPart extends BaseClientSideWebPart<IStoragePulseWebPartProps> {
  private themeVariant: IReadonlyTheme | undefined;
  private teamsTheme: string | undefined;

  protected async onInit(): Promise<void> {
    // Section backgrounds (and the page theme) arrive as a theme variant; a
    // dark section gets the dark look automatically.
    const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this.themeVariant = themeProvider.tryGetTheme();
    themeProvider.themeChangedEvent.add(this, (args: ThemeChangedEventArgs) => {
      this.themeVariant = args.theme;
      this.render();
    });

    // In Microsoft Teams the host's theme (default, dark, contrast) decides.
    const teams = this.context.sdks && this.context.sdks.microsoftTeams;
    if (teams) {
      try {
        const teamsContext = await teams.teamsJs.app.getContext();
        this.teamsTheme = teamsContext.app.theme;
        teams.teamsJs.app.registerOnThemeChangeHandler((theme: string) => {
          this.teamsTheme = theme;
          this.render();
        });
      } catch {
        // Not critical: fall back to the SharePoint theme.
      }
    }
    return super.onInit();
  }

  private get theme(): IStoragePulseTheme {
    const semantic = this.themeVariant && this.themeVariant.semanticColors;
    const palette = this.themeVariant && this.themeVariant.palette;
    const background = semantic ? semantic.bodyBackground : undefined;
    const light = luminance(background);
    let isDark = light !== undefined && light < 0.35;
    if (this.teamsTheme) {
      isDark = this.teamsTheme === 'dark' || this.teamsTheme === 'contrast';
    }
    return {
      isDark,
      accent: palette && palette.themePrimary ? palette.themePrimary : undefined,
      text: semantic ? semantic.bodyText : undefined
    };
  }

  public render(): void {
    const threshold = Number(this.properties.thresholdMonths);
    const staleDays = Number(this.properties.staleAfterDays);
    const element: React.ReactElement<IStoragePulseProps> = React.createElement(StoragePulse, {
      title: this.properties.title,
      scope: this.properties.scope === 'currentWeb' ? 'currentWeb' : 'siteCollection',
      thresholdMonths: THRESHOLD_OPTIONS.indexOf(threshold) >= 0 ? threshold : DEFAULT_THRESHOLD_MONTHS,
      scanPermission: this.properties.scanPermission === 'everyone' ? 'everyone' : 'owners',
      // Pages added before v2.1 have no scan mode set; they get the quick scan too.
      scanMode: this.properties.scanMode === 'detailed' ? 'detailed' : 'quick',
      // Pages added before v2.2 have no speed set; they get the gentle default too.
      scanSpeed:
        this.properties.scanSpeed === 'fast' || this.properties.scanSpeed === 'balanced' ? this.properties.scanSpeed : 'gentle',
      includeHidden: this.properties.includeHidden === true,
      excludeSystemLibraries: this.properties.excludeSystemLibraries === true,
      excludedLibraries: (this.properties.excludedLibraries || '')
        .split(/[\n;,]/)
        .map((n) => n.trim())
        .filter((n) => n.length > 0),
      staleAfterDays: isFinite(staleDays) && staleDays > 0 ? staleDays : 0,
      theme: this.theme,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneDropdown('thresholdMonths', {
                  label: strings.ThresholdFieldLabel,
                  options: THRESHOLD_OPTIONS.map((m) => ({ key: m, text: thresholdLabel(m) }))
                }),
                PropertyPaneSlider('staleAfterDays', {
                  label: strings.StaleAfterLabel,
                  min: 0,
                  max: 365,
                  step: 1
                })
              ]
            },
            {
              groupName: strings.ScanGroupName,
              groupFields: [
                PropertyPaneDropdown('scope', {
                  label: strings.ScopeFieldLabel,
                  options: [
                    { key: 'siteCollection', text: strings.ScopeSiteCollection },
                    { key: 'currentWeb', text: strings.ScopeCurrentWeb }
                  ]
                }),
                PropertyPaneDropdown('scanMode', {
                  label: strings.ScanModeLabel,
                  options: [
                    { key: 'quick', text: strings.ScanModeQuick },
                    { key: 'detailed', text: strings.ScanModeDetailed }
                  ]
                }),
                PropertyPaneDropdown('scanSpeed', {
                  label: strings.ScanSpeedLabel,
                  options: [
                    { key: 'gentle', text: strings.ScanSpeedGentle },
                    { key: 'balanced', text: strings.ScanSpeedBalanced },
                    { key: 'fast', text: strings.ScanSpeedFast }
                  ]
                }),
                PropertyPaneToggle('includeHidden', {
                  label: strings.IncludeHiddenLabel
                }),
                PropertyPaneToggle('excludeSystemLibraries', {
                  label: strings.ExcludeSystemLabel
                }),
                PropertyPaneTextField('excludedLibraries', {
                  label: strings.ExcludedLibrariesLabel,
                  description: strings.ExcludedLibrariesDescription,
                  multiline: true,
                  rows: 4
                })
              ]
            },
            {
              groupName: strings.AccessGroupName,
              groupFields: [
                PropertyPaneDropdown('scanPermission', {
                  label: strings.ScanPermissionLabel,
                  options: [
                    { key: 'owners', text: strings.ScanPermissionOwners },
                    { key: 'everyone', text: strings.ScanPermissionEveryone }
                  ]
                })
              ]
            },
            {
              groupName: strings.AboutGroupName,
              groupFields: [
                PropertyPaneLabel('about', {
                  text: format(strings.AboutText, { version: this.manifest.version })
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
