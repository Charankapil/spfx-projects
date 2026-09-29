import { SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * SharePoint puts the useful part of a failure in the response body, so read
 * it rather than reporting a bare status code - "Request failed (500)" says
 * nothing, while the body names the actual managed property or syntax it
 * choked on.
 */
export async function describeError(response: SPHttpClientResponse): Promise<string> {
  try {
    const body = await response.text();
    if (!body) {
      return response.statusText || 'no error details returned';
    }
    try {
      const parsed = JSON.parse(body);
      const message = parsed?.error?.message;
      if (typeof message === 'string') {
        return message;
      }
      if (message && typeof message.value === 'string') {
        return message.value;
      }
    } catch {
      // Body was not JSON - fall through and surface the raw text instead.
    }
    return body.substring(0, 300);
  } catch {
    return response.statusText || 'no error details returned';
  }
}
