// Minimal stand-ins for the SPFx runtime modules the web part imports, so the
// code can run under Node and in a plain browser page.
export const SPHttpClient = { configurations: { v1: { name: 'v1' } } };
