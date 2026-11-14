import { createHash, createPrivateKey, createSign, KeyObject, randomUUID, X509Certificate } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { Http, HttpError } from './http.js';

/**
 * App-only sign-in to Entra ID (client credentials) without any dependency.
 *
 * SharePoint refuses app-only tokens obtained with a client secret, so the app proves itself
 * with a signed assertion instead:
 *  - certificate: a JWT signed with the app's private key (the classic, proven route), or
 *  - federated:   a short-lived token from the platform the job runs on (GitHub Actions OIDC,
 *                 or any workload-identity token file). Nothing secret is stored anywhere.
 */

const b64url = (b: Buffer | string): string => Buffer.from(b).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');

export interface Credential {
  readonly kind: 'certificate' | 'federated';
  /** A client assertion for the given token endpoint. */
  assertion(tokenEndpoint: string): Promise<string>;
}

/** Accepts PEM text, PEM with literal "\n", or the whole PEM base64-encoded (handy for secret stores). */
export function normalisePem(input: string): string {
  let s = (input || '').trim();
  if (!s) {
    return '';
  }
  if (s.indexOf('-----BEGIN') < 0) {
    try {
      const decoded = Buffer.from(s.replace(/\s+/g, ''), 'base64').toString('utf8');
      if (decoded.indexOf('-----BEGIN') >= 0) {
        s = decoded;
      }
    } catch {
      /* leave as is */
    }
  }
  return s.replace(/\\n/g, '\n').replace(/\r\n/g, '\n');
}

export function pemBlocks(pem: string, label: RegExp): string[] {
  const out: string[] = [];
  const re = /-----BEGIN ([A-Z0-9 ]+)-----[\s\S]*?-----END \1-----/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(pem))) {
    if (label.test(m[1])) {
      out.push(m[0]);
    }
  }
  return out;
}

export class CertificateCredential implements Credential {
  public readonly kind = 'certificate' as const;
  public readonly thumbprint: string;
  public readonly expires: Date;
  public readonly subject: string;
  private readonly key: KeyObject;
  private readonly x5t: string;

  /**
   * @param pem  the certificate and its private key, in one PEM text or two (certificate first or second).
   */
  constructor(private clientId: string, pem: string, passphrase?: string) {
    const text = normalisePem(pem);
    const certs = pemBlocks(text, /^CERTIFICATE$/);
    const keys = pemBlocks(text, /PRIVATE KEY$/);
    if (!certs.length) {
      throw new Error('No certificate found. Provide the PEM text that contains "-----BEGIN CERTIFICATE-----".');
    }
    if (!keys.length) {
      throw new Error('No private key found. Provide the PEM text that contains "-----BEGIN PRIVATE KEY-----" (see setup/README).');
    }
    let cert: X509Certificate;
    try {
      cert = new X509Certificate(certs[0]);
    } catch (e) {
      throw new Error(`The certificate could not be read: ${(e as Error).message}`);
    }
    try {
      this.key = createPrivateKey({ key: keys[0], passphrase });
    } catch (e) {
      throw new Error(`The private key could not be read${/ENCRYPTED/.test(keys[0]) ? ' (it is encrypted: set ACCESS_REPORT_CERTIFICATE_PASSWORD or export it without a password)' : ''}: ${(e as Error).message}`);
    }
    if (this.key.asymmetricKeyType !== 'rsa') {
      throw new Error(`The private key is ${this.key.asymmetricKeyType}; an RSA key is required.`);
    }
    if (!cert.checkPrivateKey(this.key)) {
      throw new Error('The private key does not belong to the certificate.');
    }
    const sha1 = createHash('sha1').update(cert.raw).digest();
    this.x5t = b64url(sha1);
    this.thumbprint = sha1.toString('hex').toUpperCase();
    this.expires = new Date(cert.validTo);
    this.subject = cert.subject.replace(/\n/g, ', ');
  }

  public daysLeft(now = Date.now()): number {
    return Math.floor((this.expires.getTime() - now) / 86400000);
  }

  public async assertion(tokenEndpoint: string): Promise<string> {
    const now = Math.floor(Date.now() / 1000);
    const header = { alg: 'RS256', typ: 'JWT', x5t: this.x5t };
    const payload = { aud: tokenEndpoint, iss: this.clientId, sub: this.clientId, jti: randomUUID(), nbf: now - 60, iat: now, exp: now + 540 };
    const unsigned = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(payload))}`;
    const signature = createSign('RSA-SHA256').update(unsigned).sign(this.key);
    return `${unsigned}.${b64url(signature)}`;
  }
}

export interface FederatedSource {
  /** GitHub Actions: the URL and bearer token the runner provides when the job has "id-token: write". */
  requestUrl?: string;
  requestToken?: string;
  /** Any other platform: a file that holds the current token (AZURE_FEDERATED_TOKEN_FILE). */
  tokenFile?: string;
  audience?: string;
}

export class FederatedCredential implements Credential {
  public readonly kind = 'federated' as const;

  constructor(private http: Http, private source: FederatedSource) {
    if (!(source.requestUrl && source.requestToken) && !source.tokenFile) {
      throw new Error(
        'Federated sign-in needs a token source. In GitHub Actions give the job "permissions: id-token: write"; elsewhere set AZURE_FEDERATED_TOKEN_FILE.'
      );
    }
  }

  public async assertion(): Promise<string> {
    if (this.source.tokenFile) {
      const t = (await readFile(this.source.tokenFile, 'utf8')).trim();
      if (!t) {
        throw new Error('The federated token file is empty.');
      }
      return t;
    }
    const url = new URL(this.source.requestUrl as string);
    url.searchParams.set('audience', this.source.audience || 'api://AzureADTokenExchange');
    const r = await this.http.request<{ value?: string }>(url.toString(), { headers: { Authorization: `Bearer ${this.source.requestToken}`, Accept: 'application/json' } });
    if (!r.data || !r.data.value) {
      throw new Error('The platform did not return an identity token.');
    }
    return r.data.value;
  }
}

/** What went wrong at sign-in, in words an admin can act on. */
export function explainSignInError(e: unknown, credential: Credential['kind']): string {
  const raw = e instanceof HttpError ? `${e.message} ${e.body}` : (e as Error).message || String(e);
  const code = (/AADSTS\d+/.exec(raw) || [''])[0];
  const hints: Record<string, string> = {
    AADSTS700027: 'The certificate is not registered on the app registration (or it has expired). Upload the .cer under Certificates & secrets.',
    AADSTS700016: 'No app registration with this client id exists in this tenant. Check ACCESS_REPORT_CLIENT_ID and ACCESS_REPORT_TENANT_ID.',
    AADSTS90002: 'The tenant was not found. Check ACCESS_REPORT_TENANT_ID.',
    AADSTS900023: 'The tenant id is not valid. Check ACCESS_REPORT_TENANT_ID.',
    AADSTS500011: 'The resource was not found in this tenant. Check that the report site address belongs to this tenant.',
    AADSTS70021: 'No federated credential on the app registration matches this job. Check its issuer, subject (repository and branch / environment) and audience.',
    AADSTS700213: 'No federated credential on the app registration matches this job. Check its issuer, subject (repository and branch / environment) and audience.',
    AADSTS7000215: 'A client secret was rejected. This tool signs in with a certificate or a federated credential, not a secret.',
    AADSTS700024: 'The sign-in assertion is outside its valid time. Check the clock of the machine running the job.',
    AADSTS53003: 'A Conditional Access policy blocked the sign-in. Exclude this app (workload identity) or allow the runner location.'
  };
  const hint = hints[code] || (credential === 'federated' ? 'Check the federated credential on the app registration.' : 'Check the client id, tenant id and certificate.');
  const first = raw.replace(/\s+/g, ' ').substring(0, 300);
  return `Sign-in failed${code ? ` (${code})` : ''}. ${hint} Details: ${first}`;
}

interface CachedToken {
  token: string;
  expires: number;
}

export class TokenProvider {
  private cache = new Map<string, CachedToken>();
  private pending = new Map<string, Promise<string>>();
  public readonly tokenEndpoint: string;

  constructor(private http: Http, loginBase: string, tenantId: string, private clientId: string, public readonly credential: Credential) {
    this.tokenEndpoint = `${loginBase.replace(/\/+$/, '')}/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  }

  /** A bearer token for a resource, e.g. https://graph.microsoft.com or https://contoso.sharepoint.com. */
  public async getToken(resource: string, force = false): Promise<string> {
    const key = resource.replace(/\/+$/, '').toLowerCase();
    const hit = this.cache.get(key);
    if (!force && hit && hit.expires - Date.now() > 5 * 60000) {
      return hit.token;
    }
    // One sign-in at a time per resource, however many requests are waiting for it.
    let p = this.pending.get(key);
    if (!p) {
      p = this.fetchToken(key).finally(() => this.pending.delete(key));
      this.pending.set(key, p);
    }
    return p;
  }

  private async fetchToken(resource: string): Promise<string> {
    try {
      const body = new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: this.clientId,
        scope: `${resource}/.default`,
        client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
        client_assertion: await this.credential.assertion(this.tokenEndpoint)
      }).toString();
      const r = await this.http.request<{ access_token?: string; expires_in?: number }>(this.tokenEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
        body
      });
      if (!r.data || !r.data.access_token) {
        throw new Error('The sign-in service returned no token.');
      }
      this.cache.set(resource, { token: r.data.access_token, expires: Date.now() + (Number(r.data.expires_in) || 3600) * 1000 });
      return r.data.access_token;
    } catch (e) {
      throw new Error(explainSignInError(e, this.credential.kind));
    }
  }
}

/** The application permissions (roles) inside an access token; [] when it cannot be read. */
export function tokenRoles(token: string): string[] {
  try {
    const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString('utf8')) as { roles?: unknown };
    return Array.isArray(payload.roles) ? payload.roles.map(String) : [];
  } catch {
    return [];
  }
}
