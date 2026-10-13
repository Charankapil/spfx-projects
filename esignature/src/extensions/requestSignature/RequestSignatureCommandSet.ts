import { BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';

/**
 * "Request signatures" in the document library command bar. Shown when one
 * PDF or Word file is selected; opens the eSignature page with that file
 * pre-selected in the send wizard.
 */
export interface IRequestSignatureCommandSetProperties {
  /** Page hosting the eSignature web part: absolute, or relative to the current site. */
  pageUrl: string;
}

const SUPPORTED = /\.(pdf|docx?|docm|dotx|odt|rtf)$/i;

export default class RequestSignatureCommandSet extends BaseListViewCommandSet<IRequestSignatureCommandSetProperties> {
  public onInit(): Promise<void> {
    const command: Command = this.tryGetCommand('REQUEST_SIGNATURE');
    if (command) {
      command.visible = false;
    }
    this.context.listView.listViewStateChangedEvent.add(this, this.onStateChanged);
    return Promise.resolve();
  }

  private onStateChanged = (): void => {
    const command: Command = this.tryGetCommand('REQUEST_SIGNATURE');
    const rows = this.context.listView.selectedRows || [];
    const name = rows.length === 1 ? String(rows[0].getValueByName('FileLeafRef') || '') : '';
    const visible = rows.length === 1 && SUPPORTED.test(name);
    if (command && command.visible !== visible) {
      command.visible = visible;
      this.raiseOnChange();
    }
  };

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== 'REQUEST_SIGNATURE' || event.selectedRows.length !== 1) {
      return;
    }
    const fileRef = String(event.selectedRows[0].getValueByName('FileRef'));
    const configured = (this.properties.pageUrl || 'SitePages/eSignature.aspx').trim();
    const page = /^https?:\/\//i.test(configured)
      ? configured
      : `${this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')}/${configured.replace(/^\//, '')}`;
    window.location.href = `${page}${page.indexOf('?') >= 0 ? '&' : '?'}esignSource=${encodeURIComponent(fileRef)}`;
  }
}
