import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'GovernanceScorecardWebPartStrings';

import { GovernanceScorecard, IGovernanceScorecardProps } from './components/GovernanceScorecard';
import { ScorecardScanner } from './services/ScorecardScanner';
import { SnapshotStore } from './services/SnapshotStore';

export interface IGovernanceScorecardWebPartProps {
  title: string;
  requiredFields: string;
  staleYears: number;
  maxDocsPerSite: number;
  configSiteUrl: string;
  navListTitle: string;
  snapshotListTitle: string;
}

function lines(text: string): string[] {
  return (text || '')
    .split(/\r?\n|;/)
    .map((l) => l.trim())
    .filter(Boolean);
}

export default class GovernanceScorecardWebPart extends BaseClientSideWebPart<IGovernanceScorecardWebPartProps> {
  /** Kept across renders so a re-render does not reload snapshots or drop an unsaved scan. */
  private services: { scanner: ScorecardScanner; store: SnapshotStore } | undefined;

  private get hubUrl(): string {
    const origin = /^(https?:\/\/[^/]+)/i.exec(this.context.pageContext.site.absoluteUrl)?.[1] || '';
    const raw = (this.properties.configSiteUrl || '').trim();
    if (!raw) {
      return this.context.pageContext.web.absoluteUrl;
    }
    const trimmed = raw.replace(/\/+$/, '');
    if (!trimmed) {
      return origin;
    }
    return trimmed.indexOf('/') === 0 ? origin + trimmed : trimmed;
  }

  public render(): void {
    const p = this.properties;
    if (!this.services) {
      const hubUrl = this.hubUrl;
      this.services = {
        scanner: new ScorecardScanner(this.context.spHttpClient, {
          configSiteUrl: hubUrl,
          navListTitle: p.navListTitle || 'ECM Navigation',
          requiredFields: lines(p.requiredFields),
          staleYears: Math.max(1, p.staleYears || 3),
          maxDocsPerSite: Math.max(500, p.maxDocsPerSite || 5000)
        }),
        store: new SnapshotStore(this.context.spHttpClient, hubUrl, p.snapshotListTitle || 'ECM Scorecard Snapshots')
      };
    }
    const element: React.ReactElement<IGovernanceScorecardProps> = React.createElement(GovernanceScorecard, {
      title: p.title,
      scanner: this.services.scanner,
      store: this.services.store,
      staleYears: Math.max(1, p.staleYears || 3)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.services = undefined;
    this.render();
  }

  /** Changes apply with the pane's Apply button, not on every keystroke. */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
            { groupName: strings.DisplayGroup, groupFields: [PropertyPaneTextField('title', { label: strings.TitleLabel })] },
            {
              groupName: strings.MeasureGroup,
              groupFields: [
                PropertyPaneTextField('requiredFields', {
                  label: strings.RequiredFieldsLabel,
                  description: strings.RequiredFieldsDescription,
                  multiline: true,
                  rows: 6
                }),
                PropertyPaneSlider('staleYears', { label: strings.StaleYearsLabel, min: 1, max: 10, step: 1 }),
                PropertyPaneSlider('maxDocsPerSite', { label: strings.MaxDocsLabel, min: 500, max: 20000, step: 500 })
              ]
            },
            {
              groupName: strings.SourcesGroup,
              groupFields: [
                PropertyPaneTextField('configSiteUrl', { label: strings.ConfigSiteLabel, description: strings.ConfigSiteDescription }),
                PropertyPaneTextField('navListTitle', { label: strings.NavListLabel }),
                PropertyPaneTextField('snapshotListTitle', { label: strings.SnapshotListLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
