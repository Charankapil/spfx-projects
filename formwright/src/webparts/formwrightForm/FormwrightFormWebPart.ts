import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { FormMode } from '../../common/model/IFormDefinition';
import { FormStore } from '../../common/services/FormStore';
import { ListService } from '../../common/services/ListService';
import { FormPage, IFormPageProps } from './FormPage';

export interface IFormwrightFormWebPartProps {
  listId: string;
  formId: string;
  /** Query string parameter that carries the id of an item to open, e.g. ?itemId=12. */
  itemParameter: string;
  existingItemMode: FormMode;
  allowAnother: boolean;
}

export default class FormwrightFormWebPart extends BaseClientSideWebPart<IFormwrightFormWebPartProps> {
  private listOptions: IPropertyPaneDropdownOption[] = [];
  private formOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const param = (this.properties.itemParameter || 'itemId').trim();
    const raw = new URLSearchParams(window.location.search).get(param);
    const itemId = raw && /^\d+$/.test(raw) ? Number(raw) : undefined;
    const element: React.ReactElement<IFormPageProps> = React.createElement(FormPage, {
      ctx: this.context,
      listId: this.properties.listId,
      formId: this.properties.formId || undefined,
      itemId,
      mode: itemId ? this.properties.existingItemMode || 'Edit' : 'New',
      allowAnother: this.properties.allowAnother !== false,
      onConfigure: () => this.context.propertyPane.open()
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    const service = new ListService(this.context);
    try {
      const lists = await service.getLists();
      this.listOptions = lists.filter((l) => !l.isLibrary).map((l) => ({ key: l.id, text: l.title }));
    } catch {
      this.listOptions = [];
    }
    await this.loadForms();
    this.context.propertyPane.refresh();
  }

  private async loadForms(): Promise<void> {
    this.formOptions = [{ key: '', text: 'Automatic (by audience)' }];
    if (!this.properties.listId) {
      return;
    }
    try {
      const saved = await new FormStore(this.context).load(this.properties.listId);
      saved.forms.forEach((f) => this.formOptions.push({ key: f.id, text: f.name }));
    } catch {
      // Keep the automatic option only.
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'listId' && oldValue !== newValue) {
      this.properties.formId = '';
      this.loadForms()
        .then(() => this.context.propertyPane.refresh())
        .catch(() => undefined);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Show a list form on this page. Design forms with the Formwright - Designer web part.' },
          groups: [
            {
              groupName: 'Form',
              groupFields: [
                PropertyPaneDropdown('listId', { label: 'List', options: this.listOptions }),
                PropertyPaneDropdown('formId', { label: 'Form', options: this.formOptions, disabled: !this.properties.listId }),
                PropertyPaneToggle('allowAnother', { label: 'After submitting, offer to submit another response', onText: 'Yes', offText: 'No' })
              ]
            },
            {
              groupName: 'Opening existing items',
              groupFields: [
                PropertyPaneTextField('itemParameter', {
                  label: 'Query string parameter with the item id',
                  description: 'With "itemId", page.aspx?itemId=12 opens item 12 instead of a new item.'
                }),
                PropertyPaneDropdown('existingItemMode', {
                  label: 'Open existing items in',
                  options: [
                    { key: 'Edit', text: 'Edit mode' },
                    { key: 'Display', text: 'Display mode' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
