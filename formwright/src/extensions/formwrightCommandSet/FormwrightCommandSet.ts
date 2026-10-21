import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import { DesignerOverlay } from './DesignerOverlay';

/* eslint-disable-next-line @typescript-eslint/no-empty-interface */
export interface IFormwrightCommandSetProperties {}

const COMMAND = 'FORMWRIGHT_DESIGN';

/**
 * Adds "Customize forms" to the list's command bar. It opens the Formwright
 * designer over the list, already pointed at it. Only people who can manage
 * the list see the button; SharePoint enforces the same permission when the
 * forms are saved or switched on.
 */
export default class FormwrightCommandSet extends BaseListViewCommandSet<IFormwrightCommandSetProperties> {
  private host: HTMLDivElement | undefined;

  public onInit(): Promise<void> {
    this.updateCommand();
    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewStateChanged);
    return Promise.resolve();
  }

  private onListViewStateChanged(): void {
    this.updateCommand();
  }

  private updateCommand(): void {
    const command: Command | undefined = this.tryGetCommand(COMMAND);
    if (!command) {
      return;
    }
    const list = this.context.pageContext.list;
    const visible = !!list && list.permissions.hasPermission(SPPermission.manageLists);
    if (command.visible !== visible) {
      command.visible = visible;
      this.raiseOnChange();
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== COMMAND) {
      return;
    }
    const list = this.context.listView.list;
    const listId = list ? list.guid.toString() : this.context.pageContext.list ? this.context.pageContext.list.id.toString() : '';
    if (!listId) {
      return;
    }
    if (!this.host) {
      this.host = document.createElement('div');
      document.body.appendChild(this.host);
    }
    ReactDom.render(
      React.createElement(DesignerOverlay, {
        ctx: this.context,
        listId,
        listTitle: list ? list.title : '',
        onClose: () => this.close()
      }),
      this.host
    );
  }

  private close(): void {
    if (this.host) {
      ReactDom.unmountComponentAtNode(this.host);
      this.host.remove();
      this.host = undefined;
    }
  }

  protected onDispose(): void {
    this.close();
    super.onDispose();
  }
}
