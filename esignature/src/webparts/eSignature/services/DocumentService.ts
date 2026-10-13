import { SPHttpClient } from '@microsoft/sp-http';
import { odataString, SpRest } from './SpRest';

/**
 * Reading, uploading and converting files. Paths are server-relative
 * ("/sites/x/Shared Documents/a.pdf") and always passed through @aliases so
 * names with quotes, '#' or '%' survive.
 */

export interface ILibraryInfo {
  title: string;
  url: string;
}

export interface IFolderEntry {
  name: string;
  url: string;
  isFolder: boolean;
  size?: number;
  modified?: string;
}

const alias = (value: string): string => encodeURIComponent(odataString(value));

export function fileName(url: string): string {
  return decodeURIComponent(url.substring(url.lastIndexOf('/') + 1));
}

export function folderOf(url: string): string {
  return url.substring(0, url.lastIndexOf('/'));
}

export function baseName(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.substring(0, dot) : name;
}

export function extension(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.substring(dot + 1).toLowerCase() : '';
}

export const WORD_EXTENSIONS = ['doc', 'docx', 'docm', 'dot', 'dotx', 'odt', 'rtf'];
export const SUPPORTED_EXTENSIONS = ['pdf'].concat(WORD_EXTENSIONS);

/** Remove characters SharePoint does not allow in file or folder names. */
export function safeFileName(name: string): string {
  return name.replace(/["*:<>?/\\|#%]/g, '_').replace(/^\.+/, '').trim() || 'document';
}

export class DocumentService {
  constructor(private rest: SpRest, private client: SPHttpClient) {}

  public async libraries(): Promise<ILibraryInfo[]> {
    const result = await this.rest.getJson<{ value: { Title: string; RootFolder: { ServerRelativeUrl: string } }[] }>(
      `/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseTemplate eq 101 and Hidden eq false`
    );
    return result.value
      .filter(l => !/\/ESignDocuments$/i.test(l.RootFolder.ServerRelativeUrl))
      .map(l => ({ title: l.Title, url: l.RootFolder.ServerRelativeUrl }));
  }

  public async folder(url: string): Promise<IFolderEntry[]> {
    const base = `/_api/web/GetFolderByServerRelativePath(decodedurl=@f)`;
    const [folders, files] = await Promise.all([
      this.rest.getJson<{ value: { Name: string; ServerRelativeUrl: string }[] }>(
        `${base}/Folders?@f=${alias(url)}&$select=Name,ServerRelativeUrl&$orderby=Name`
      ),
      this.rest.getJson<{ value: { Name: string; ServerRelativeUrl: string; Length: string; TimeLastModified: string }[] }>(
        `${base}/Files?@f=${alias(url)}&$select=Name,ServerRelativeUrl,Length,TimeLastModified&$orderby=Name`
      )
    ]);
    return folders.value
      .filter(f => f.Name !== 'Forms')
      .map(f => ({ name: f.Name, url: f.ServerRelativeUrl, isFolder: true }))
      .concat(
        files.value
          .filter(f => SUPPORTED_EXTENSIONS.indexOf(extension(f.Name)) >= 0)
          .map(f => ({
            name: f.Name,
            url: f.ServerRelativeUrl,
            isFolder: false,
            size: parseInt(f.Length, 10),
            modified: f.TimeLastModified
          }))
      );
  }

  /** Absolute URL of the web that contains a server-relative file URL (it may be on another site). */
  public async webUrlFor(serverRelativeUrl: string): Promise<string> {
    const origin = new URL(this.rest.webUrl).origin;
    try {
      const result = await this.rest.postJson<{ value?: string }>(`${origin}/_api/SP.Web.GetWebUrlFromPageUrl(@p)?@p=${alias(origin + serverRelativeUrl)}`);
      if (result.value) {
        return result.value;
      }
    } catch {
      // Fall back to the current web below.
    }
    return this.rest.webUrl;
  }

  private restFor(webUrl: string): SpRest {
    return webUrl === this.rest.webUrl ? this.rest : new SpRest(this.client, webUrl);
  }

  public async read(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const rest = this.restFor(await this.webUrlFor(serverRelativeUrl));
    return rest.getBinary(`/_api/web/GetFileByServerRelativePath(decodedurl=@u)/$value?@u=${alias(serverRelativeUrl)}`);
  }

  /** Read a file from the current web without the cross-site lookup (eSignature's own library). */
  public readLocal(serverRelativeUrl: string): Promise<ArrayBuffer> {
    return this.rest.getBinary(`/_api/web/GetFileByServerRelativePath(decodedurl=@u)/$value?@u=${alias(serverRelativeUrl)}`);
  }

  public async ensureFolder(serverRelativeUrl: string): Promise<void> {
    await this.rest.postJson(`/_api/web/folders/addUsingPath(DecodedUrl=@f,overwrite=false)?@f=${alias(serverRelativeUrl)}`);
  }

  /**
   * Upload without overwriting: signed documents are evidence, so an
   * existing file with the same name gets a timestamped sibling instead.
   * Pass webUrl to upload to another site (e.g. next to the source file).
   */
  public async upload(folderUrl: string, name: string, content: ArrayBuffer | Uint8Array, webUrl?: string): Promise<string> {
    const rest = webUrl ? this.restFor(webUrl) : this.rest;
    try {
      return await this.uploadRaw(rest, folderUrl, name, content);
    } catch (e) {
      if (!/exists/i.test((e as Error).message)) {
        throw e;
      }
      const stamped = `${baseName(name)} ${new Date().toISOString().replace(/[:.]/g, '-')}.${extension(name)}`;
      return this.uploadRaw(rest, folderUrl, stamped, content);
    }
  }

  private async uploadRaw(rest: SpRest, folderUrl: string, name: string, content: ArrayBuffer | Uint8Array): Promise<string> {
    const response = await rest.fetch(
      `/_api/web/GetFolderByServerRelativePath(decodedurl=@f)/Files/AddUsingPath(decodedurl=@n,overwrite=false)?@f=${alias(folderUrl)}&@n=${alias(name)}`,
      'POST',
      { body: content as ArrayBuffer }
    );
    const file: { ServerRelativeUrl: string } = await response.json();
    return file.ServerRelativeUrl;
  }

  /**
   * Converts a Word document to PDF with SharePoint's own file conversion
   * (the same service behind "Download as PDF"), via the v2.0 drive API on
   * the file's site. Returns the PDF bytes.
   */
  public async convertToPdf(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const webUrl = await this.webUrlFor(serverRelativeUrl);
    const rest = this.restFor(webUrl);
    const file = `/_api/web/GetFileByServerRelativePath(decodedurl=@u)`;
    const [site, web, item, list] = await Promise.all([
      rest.getJson<{ Id: string }>('/_api/site?$select=Id'),
      rest.getJson<{ Id: string }>('/_api/web?$select=Id'),
      rest.getJson<{ Id: number }>(`${file}/ListItemAllFields?@u=${alias(serverRelativeUrl)}&$select=Id`),
      rest.getJson<{ Id: string }>(`${file}/ListItemAllFields/ParentList?@u=${alias(serverRelativeUrl)}&$select=Id`)
    ]);
    const host = new URL(webUrl).hostname;
    return rest.getBinary(
      `/_api/v2.0/sites/${host},${site.Id},${web.Id}/lists/${list.Id}/items/${item.Id}/driveItem/content?format=pdf`
    );
  }

  /** Convert an uploaded Word file: park it in the request folder, convert, then return the PDF. */
  public async convertUploadToPdf(folderUrl: string, name: string, content: ArrayBuffer): Promise<ArrayBuffer> {
    const url = await this.upload(folderUrl, name, content);
    return this.convertToPdf(url);
  }
}
