import { SpRequestError, SpRest } from './SpRest';

/**
 * Creates the lists and the library eSignature keeps its data in, on the
 * site that hosts the web part. Run once by a site owner from the Setup
 * screen; safe to re-run (anything that already exists is left alone).
 *
 * Requests and Activity are created with "Create items and edit items that
 * were created by the user": only the sender can change their request, and
 * only a signer can change their own activity. That, plus SharePoint's own
 * Author stamp, is what stops one person recording a signature for another.
 */

export const STORE = {
  requests: 'Lists/ESignRequests',
  activity: 'Lists/ESignActivity',
  templates: 'Lists/ESignTemplates',
  library: 'ESignDocuments'
};

interface IListSpec {
  url: string;
  internalName: string;
  title: string;
  description: string;
  baseTemplate: number;
  fields: string[];
  ownItemsOnly: boolean;
}

const note = (name: string, display: string): string =>
  `<Field Type="Note" Name="${name}" StaticName="${name}" DisplayName="${display}" NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`;
const text = (name: string, display: string, indexed = false): string =>
  `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${display}" MaxLength="255"${indexed ? ' Indexed="TRUE"' : ''} />`;

const LISTS: IListSpec[] = [
  {
    url: STORE.requests,
    internalName: 'ESignRequests',
    title: 'eSignature Requests',
    description: 'Signature requests sent from the eSignature web part. Do not edit items here directly.',
    baseTemplate: 100,
    fields: [text('RequestId', 'Request ID', true), text('SenderStatus', 'Sender status'), note('Payload', 'Payload')],
    ownItemsOnly: true
  },
  {
    url: STORE.activity,
    internalName: 'ESignActivity',
    title: 'eSignature Activity',
    description: 'Audit trail of every eSignature event. Items cannot be changed without invalidating them.',
    baseTemplate: 100,
    fields: [
      text('RequestId', 'Request ID', true),
      text('EventType', 'Event'),
      text('Slot', 'Recipient slot'),
      text('DocHash', 'Document SHA-256', true),
      note('Payload', 'Payload')
    ],
    ownItemsOnly: true
  },
  {
    url: STORE.templates,
    internalName: 'ESignTemplates',
    title: 'eSignature Templates',
    description: 'Reusable recipient roles and field layouts for eSignature requests.',
    baseTemplate: 100,
    fields: [note('Payload', 'Payload')],
    ownItemsOnly: false
  },
  {
    url: STORE.library,
    internalName: 'ESignDocuments',
    title: 'eSignature Documents',
    description: 'Documents sent for signature and their signed copies, one folder per request.',
    baseTemplate: 101,
    fields: [],
    ownItemsOnly: false
  }
];

export class Provisioning {
  constructor(private rest: SpRest, private webServerRelativeUrl: string) {}

  private serverRelative(url: string): string {
    return `${this.webServerRelativeUrl.replace(/\/$/, '')}/${url}`;
  }

  public listApi(url: string): string {
    return `/_api/web/GetList(@l)?@l='${encodeURIComponent(this.serverRelative(url).replace(/'/g, "''"))}'`;
  }

  /** `/_api/web/GetList(@l)` with a sub-path inserted before the alias. */
  public listPath(url: string, path: string, query = ''): string {
    const alias = `@l='${encodeURIComponent(this.serverRelative(url).replace(/'/g, "''"))}'`;
    return `/_api/web/GetList(@l)/${path}?${alias}${query ? `&${query}` : ''}`;
  }

  public async isProvisioned(): Promise<boolean> {
    for (const spec of LISTS) {
      if (!(await this.exists(spec.url))) {
        return false;
      }
    }
    return true;
  }

  private async exists(url: string): Promise<boolean> {
    try {
      const list = await this.rest.tryGetJson<{ Id: string }>(`${this.listApi(url)}&$select=Id`);
      return !!list;
    } catch (e) {
      // GetList answers a missing list with 404 on some tenants and 500 (FileNotFound) on others.
      if (e instanceof SpRequestError && (e.status === 404 || e.status === 500)) {
        return false;
      }
      throw e;
    }
  }

  public async provision(log: (message: string) => void): Promise<void> {
    for (const spec of LISTS) {
      if (await this.exists(spec.url)) {
        log(`${spec.title}: already exists.`);
      } else {
        log(`${spec.title}: creating...`);
        await this.rest.postJson('/_api/web/lists', {
          Title: spec.internalName,
          Description: spec.description,
          BaseTemplate: spec.baseTemplate,
          AllowContentTypes: false,
          ContentTypesEnabled: false,
          EnableVersioning: true
        });
        // Created under the internal name so the URL is stable; now give it a friendly title.
        await this.rest.merge(this.listApi(spec.url), { Title: spec.title, OnQuickLaunch: false });
      }

      const existing = await this.rest.getJson<{ value: { InternalName: string }[] }>(
        this.listPath(spec.url, 'fields', '$select=InternalName')
      );
      const names = existing.value.map(f => f.InternalName);
      for (const schema of spec.fields) {
        const name = (/Name="([^"]+)"/.exec(schema) || [])[1];
        if (names.indexOf(name) < 0) {
          log(`${spec.title}: adding column ${name}.`);
          await this.rest.postJson(this.listPath(spec.url, 'fields/CreateFieldAsXml'), {
            parameters: { SchemaXml: schema, Options: 8 }
          });
        }
      }

      if (spec.ownItemsOnly) {
        await this.restrictEditsToAuthors(spec, log);
      }
    }
    log('Setup complete.');
  }

  private async restrictEditsToAuthors(spec: IListSpec, log: (message: string) => void): Promise<void> {
    try {
      await this.rest.merge(this.listApi(spec.url), { WriteSecurity: 2 });
      log(`${spec.title}: people can only edit items they created.`);
      return;
    } catch {
      // Older REST surfaces reject WriteSecurity - fall back to CSOM below.
    }
    const list = await this.rest.getJson<{ Id: string }>(`${this.listApi(spec.url)}&$select=Id`);
    const xml =
      '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="eSignature" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">' +
      '<Actions>' +
      '<SetProperty Id="1" ObjectPathId="13" Name="WriteSecurity"><Parameter Type="Number">2</Parameter></SetProperty>' +
      '<Method Name="Update" Id="2" ObjectPathId="13" />' +
      '</Actions>' +
      '<ObjectPaths>' +
      '<StaticProperty Id="10" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />' +
      '<Property Id="11" ParentId="10" Name="Web" />' +
      '<Property Id="12" ParentId="11" Name="Lists" />' +
      `<Method Id="13" ParentId="12" Name="GetById"><Parameters><Parameter Type="Guid">{${list.Id}}</Parameter></Parameters></Method>` +
      '</ObjectPaths>' +
      '</Request>';
    const response = await this.rest.fetch('/_vti_bin/client.svc/ProcessQuery', 'POST', {
      headers: { 'Content-Type': 'text/xml' },
      body: xml
    });
    const result = await response.json();
    if (Array.isArray(result) && result[0] && result[0].ErrorInfo) {
      log(`${spec.title}: WARNING - could not restrict edits to authors (${result[0].ErrorInfo.ErrorMessage}). Set it manually in List settings > Advanced settings > Item-level permissions.`);
    } else {
      log(`${spec.title}: people can only edit items they created.`);
    }
  }
}
