import { FormMode } from '../model/IFormDefinition';
import { getJson, guidOnly, ISpContext, postJson, quoteForUrl } from './http';

/** Component id of the Form Studio form customizer (see FormStudioFormCustomizer.manifest.json). */
export const FORM_CUSTOMIZER_ID = '66367e94-d452-47cf-9d7f-cb730d61bafd';

const PROPERTY: { [mode in FormMode]: string } = {
  New: 'NewFormClientSideComponentId',
  Edit: 'EditFormClientSideComponentId',
  Display: 'DisplayFormClientSideComponentId'
};

export interface IContentTypeBinding {
  id: string;
  name: string;
  modes: { [mode in FormMode]: boolean };
}

/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * SharePoint decides which form a list item opens with per content type:
 * when New/Edit/DisplayFormClientSideComponentId names a form customizer,
 * the modern UI loads it instead of the built-in form. This is the supported
 * replacement for the "Customize forms" (Power Apps) switch, and is what
 * PnP PowerShell's Set-PnPContentType -NewFormClientSideComponentId sets.
 */
export class FormDeployment {
  constructor(private ctx: ISpContext, private webUrl: string = ctx.pageContext.web.absoluteUrl) {}

  private list(listId: string): string {
    return `${this.webUrl}/_api/web/lists(guid'${guidOnly(listId)}')`;
  }

  public async getBindings(listId: string): Promise<IContentTypeBinding[]> {
    const json = await getJson<{ value: any[] }>(
      this.ctx,
      `${this.list(listId)}/contenttypes?$select=StringId,Name,Hidden,NewFormClientSideComponentId,EditFormClientSideComponentId,DisplayFormClientSideComponentId`
    );
    const is = (value: string | null | undefined): boolean => !!value && guidOnly(value).toLowerCase() === FORM_CUSTOMIZER_ID;
    return json.value
      // Folders keep SharePoint's own form.
      .filter((ct) => !ct.Hidden && String(ct.StringId).indexOf('0x0120') !== 0)
      .map((ct) => ({
        id: ct.StringId,
        name: ct.Name,
        modes: {
          New: is(ct.NewFormClientSideComponentId),
          Edit: is(ct.EditFormClientSideComponentId),
          Display: is(ct.DisplayFormClientSideComponentId)
        }
      }));
  }

  /**
   * Points the chosen modes of every content type at Form Studio (enabled)
   * or back at SharePoint's default form (disabled).
   */
  public async apply(listId: string, modes: FormMode[], enabled: boolean): Promise<void> {
    const bindings = await this.getBindings(listId);
    for (const ct of bindings) {
      const body: { [property: string]: string } = {};
      modes.forEach((m) => {
        body[PROPERTY[m]] = enabled ? FORM_CUSTOMIZER_ID : '';
      });
      await postJson(this.ctx, `${this.list(listId)}/contenttypes('${quoteForUrl(ct.id)}')`, body, {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      });
    }
  }
}

/** URL of the modern form page for a list item, which loads the customizer when one is bound. */
export function listFormUrl(webUrl: string, listId: string, mode: FormMode, itemId?: number, source?: string): string {
  const pageType = mode === 'New' ? 8 : mode === 'Edit' ? 6 : 4;
  let url = `${webUrl}/_layouts/15/SPListForm.aspx?PageType=${pageType}&List=${encodeURIComponent(`{${guidOnly(listId)}}`)}`;
  if (itemId) {
    url += `&ID=${itemId}`;
  }
  if (source) {
    url += `&Source=${encodeURIComponent(source)}`;
  }
  return url;
}
