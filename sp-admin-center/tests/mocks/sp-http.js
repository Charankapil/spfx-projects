// Minimal stand-in for @microsoft/sp-http so the real service code can run under Node.
class SPHttpClient {
  constructor(handler) {
    this.handler = handler;
  }
  get(url, config, opts) {
    return this.handler('GET', url, config, opts || {});
  }
  post(url, config, opts) {
    return this.handler('POST', url, config, opts || {});
  }
}
SPHttpClient.configurations = { v1: { id: 'v1', overrideWith: (o) => Object.assign({ id: 'v1' }, o, { v3: true }) } };
module.exports = { SPHttpClient, ODataVersion: { v3: 'v3', v4: 'v4' } };
