// Minimal stand-in for @microsoft/sp-http in the local preview harness.
export const SPHttpClient = { configurations: { v1: { overrideWith: () => ({}) } } };
export const ODataVersion = { v3: 'v3', v4: 'v4' };
