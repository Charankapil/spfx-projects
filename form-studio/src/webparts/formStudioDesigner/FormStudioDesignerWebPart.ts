import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { FormStudioDesigner, IFormStudioDesignerProps } from './components/FormStudioDesigner';

export interface IFormStudioDesignerWebPartProps {
  /** Last list designed on this page, so the designer reopens it. */
  listId: string;
}

/**
 * The designer. Best on a full-width section or a single-part app page;
 * it needs Manage Lists on the list to save.
 */
export default class FormStudioDesignerWebPart extends BaseClientSideWebPart<IFormStudioDesignerWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IFormStudioDesignerProps> = React.createElement(FormStudioDesigner, {
      ctx: this.context,
      listId: this.properties.listId || '',
      onListChange: (listId: string) => {
        this.properties.listId = listId;
      }
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
