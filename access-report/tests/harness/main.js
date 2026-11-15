// The real web part component on a fake tenant. ?s=fresh starts before set-up; anything else
// starts set up, with people and groups already chosen.
const React = require('react');
const ReactDom = require('react-dom');
const { initializeIcons } = require('@fluentui/react');
const { AccessReport } = require('../../lib/webparts/accessReport/components/AccessReport');
const { SPClient } = require('../../lib/webparts/accessReport/core/SPClient');
const { SPHttpClient } = require('../mocks/sp-http');
const { FakeTenant } = require('../fakeTenant');
const scenarios = require('../scenarios');

const fresh = /[?&]s=fresh/.test(location.search);
const tenant = fresh ? new FakeTenant({ sites: scenarios.richSites() }) : scenarios.rich();
if (fresh) {
  tenant.unknownUsers = [];
}
tenant.picker = [
  { Key: scenarios.ANN, DisplayText: 'Ann Example', EntityType: 'User', EntityData: { Email: 'ann@contoso.com' } },
  { Key: scenarios.MARKETING, DisplayText: 'Marketing', EntityType: 'FormsRole', EntityData: { Email: 'marketing@contoso.com' } },
  { Key: scenarios.AUDITORS, DisplayText: 'Finance Auditors', EntityType: 'SecGroup', EntityData: {} },
  { Key: 'IT Members', DisplayText: 'IT Members', EntityType: 'SPGroup' }
];
const handler = tenant.spHttpHandler();
const slow = async (...args) => {
  await new Promise((r) => setTimeout(r, Number((/[?&]delay=(\d+)/.exec(location.search) || [0, 15])[1])));
  return handler(...args);
};
window.__tenant = tenant;
window.__errors = [];
window.addEventListener('error', (e) => window.__errors.push(String(e.message)));
window.addEventListener('unhandledrejection', (e) => window.__errors.push(String(e.reason)));

initializeIcons();
ReactDom.render(
  React.createElement(AccessReport, {
    client: new SPClient(new SPHttpClient(slow), tenant.origin, 5),
    homeWebUrl: tenant.home,
    heading: 'Access Report',
    tenantId: '33333333-3333-3333-3333-333333333333',
    version: '1.0.0'
  }),
  document.getElementById('root')
);
