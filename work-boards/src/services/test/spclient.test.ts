import { SpClient } from '../SpClient';

/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Fake SPHttpClient that applies the same header rule as the real one (SPHttpClientHelper):
 * with jsonRequest on, a non-GET request with no Content-Type needs OData-Version 3.0 or 4.0.
 * The app sends 'odata-version': '' (no version), so every write must carry a Content-Type.
 */
class StrictHttp {
  public requests: { url: string; headers: { [k: string]: string } }[] = [];

  private check(options: any): void {
    const h: { [k: string]: string } = {};
    Object.keys(options.headers || {}).forEach(k => { h[k.toLowerCase()] = options.headers[k]; });
    const version = h['odata-version'];
    if (!h['content-type'] && version !== '3.0' && version !== '4.0') {
      throw new Error('ISPHttpClientConfiguration.jsonRequest is enabled, which requires the "OData-Version" header to be 3.0 or 4.0');
    }
  }

  private respond(): any {
    return { ok: true, status: 200, headers: { get: () => '' }, text: async () => '', json: async () => ({ value: [] }) };
  }

  public async get(url: string, _c: unknown, options: any): Promise<any> {
    this.requests.push({ url, headers: options.headers });
    return this.respond();
  }

  public async post(url: string, _c: unknown, options: any): Promise<any> {
    this.check(options);
    this.requests.push({ url, headers: options.headers });
    return this.respond();
  }
}

describe('SpClient request headers', () => {
  const make = (): { http: StrictHttp; sp: SpClient } => {
    const http = new StrictHttp();
    return { http, sp: new SpClient(http as any, {} as any, 'https://contoso.sharepoint.com/sites/team') };
  };

  it('uploads attachments with a Content-Type, so SPHttpClient accepts them', async () => {
    const { http, sp } = make();
    await sp.postBinary("web/lists(guid'x')/items(1)/AttachmentFiles/add(FileName='a.xlsx')", new ArrayBuffer(4));
    expect(http.requests[0].headers['Content-Type']).toBe('application/octet-stream');
  });

  it('sends a Content-Type on every other kind of write', async () => {
    const { sp } = make();
    await sp.post('web/lists', { Title: 'x' });
    await sp.post("web/lists(guid'x')/recycle");
    await sp.merge("web/lists(guid'x')", { Title: 'y' });
    await sp.remove("web/lists(guid'x')/items(1)");
  });
});
