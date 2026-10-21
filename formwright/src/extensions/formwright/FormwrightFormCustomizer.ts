import * as React from 'react';
import * as ReactDom from 'react-dom';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

import { FormHost } from '../../common/components/renderer/FormHost';
import { FormMode } from '../../common/model/IFormDefinition';
import { listFormUrl } from '../../common/services/FormDeployment';

/* eslint-disable-next-line @typescript-eslint/no-empty-interface */
export interface IFormwrightFormCustomizerProperties {}

/**
 * Replaces a list's New / Edit / Display forms. Which list forms use it is
 * set per content type (the designer's "Use on list" button, or PnP
 * PowerShell); the form itself comes from the definitions saved by the
 * Formwright designer for this list.
 */
export default class FormwrightFormCustomizer extends BaseFormCustomizer<IFormwrightFormCustomizerProperties> {
  private get mode(): FormMode {
    switch (this.displayMode) {
      case FormDisplayMode.New:
        return 'New';
      case FormDisplayMode.Edit:
        return 'Edit';
      default:
        return 'Display';
    }
  }

  public render(): void {
    const ctx = this.context;
    const listId = ctx.list.guid.toString();
    const element = React.createElement(FormHost, {
      ctx,
      listId,
      mode: this.mode,
      itemId: ctx.itemId,
      contentTypeId: ctx.contentType ? ctx.contentType.id : undefined,
      folderPath: ctx.folderInfo ? ctx.folderInfo.folderPath : undefined,
      onSaved: (_itemId: number, _values: unknown, redirectUrl?: string) => {
        if (redirectUrl) {
          window.location.href = redirectUrl;
        } else {
          this.formSaved();
        }
      },
      onCancel: () => this.formClosed(),
      onEdit: ctx.itemId
        ? () => {
            const source = new URLSearchParams(window.location.search).get('Source') || undefined;
            window.location.href = listFormUrl(ctx.pageContext.web.absoluteUrl, listId, 'Edit', ctx.itemId, source);
          }
        : undefined
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }
}
