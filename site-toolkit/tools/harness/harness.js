/**
 * Test page: mounts one or more Site Toolkit components against the mock
 * SharePoint. Driven by tools/harness/run.js; see that file for the checks.
 */
'use strict';

const React = require('react');
const ReactDom = require('react-dom');
const { initializeIcons } = require('@fluentui/react');

const { createMock, ORIGIN, WEB } = require('./mock-sharepoint');

const COMPONENTS = {
  StorageHogFinder: () => require('../../lib/webparts/storageHogFinder/StorageHogFinder').StorageHogFinder,
  ReadAcknowledge: () => require('../../lib/webparts/readAcknowledge/ReadAcknowledge').ReadAcknowledge,
  StaleContentRadar: () => require('../../lib/webparts/staleContentRadar/StaleContentRadar').StaleContentRadar,
  RecycleBinExplorer: () => require('../../lib/webparts/recycleBinExplorer/RecycleBinExplorer').RecycleBinExplorer,
  BulkMetadataEditor: () => require('../../lib/webparts/bulkMetadataEditor/BulkMetadataEditor').BulkMetadataEditor,
  DuplicateFileDetector: () => require('../../lib/webparts/duplicateFileDetector/DuplicateFileDetector').DuplicateFileDetector,
  MyFiles: () => require('../../lib/webparts/myFiles/MyFiles').MyFiles,
  PulsePoll: () => require('../../lib/webparts/pulsePoll/PulsePoll').PulsePoll,
  TemplateGallery: () => require('../../lib/webparts/templateGallery/TemplateGallery').TemplateGallery,
  OrgChart: () => require('../../lib/webparts/orgChart/OrgChart').OrgChart,
  FaqAccordion: () => require('../../lib/webparts/faqAccordion/FaqAccordion').FaqAccordion,
  WhoHasAccess: () => require('../../lib/webparts/whoHasAccess/WhoHasAccess').WhoHasAccess,
  BrokenLinkChecker: () => require('../../lib/webparts/brokenLinkChecker/BrokenLinkChecker').BrokenLinkChecker,
  AlertBannerSettings: () => require('../../lib/webparts/alertBannerSettings/AlertBannerSettings').AlertBannerSettings
};

const DEFAULT_PROPS = {
  ReadAcknowledge: { libraryTitle: 'Policies' },
  StaleContentRadar: { months: '12' },
  MyFiles: { rowCount: '10' },
  PulsePoll: { question: 'How was your week?', options: 'Great\nGood\nOkay\nNot good', pollKey: 'weekly' },
  TemplateGallery: { libraryTitle: 'Templates' },
  BrokenLinkChecker: { maxLinks: '300' },
  OrgChart: { managerLevels: '2', showPhotos: true }
};

const query = new URLSearchParams(window.location.search);
const options = JSON.parse(query.get('options') || '{}');
options.role = query.get('role') || 'member';
options.state = query.get('state') || 'seeded';
options.userId = Number(query.get('user') || 7);

window.__xss = false;
window.alert = () => {
  window.__xss = true;
};

const mock = createMock(options);
window.__sp = mock;

const users = { 7: ['Alex Reader', 'alex@contoso.com'], 8: ['Priya Shah', 'priya@contoso.com'] };
const user = users[options.userId] || users[7];
const pageContext = {
  web: {
    absoluteUrl: `${ORIGIN}${WEB}`,
    serverRelativeUrl: WEB,
    id: { toString: () => 'web-guid-1' },
    permissions: { hasPermission: () => options.role === 'owner' }
  },
  site: { absoluteUrl: `${ORIGIN}${WEB}`, serverRelativeUrl: WEB },
  user: { displayName: user[0], email: user[1], loginName: user[1] },
  legacyPageContext: { userId: options.userId, isSiteAdmin: false }
};

// Link checks use fetch directly (through the same queue); record them alongside the REST calls.
const realFetch = window.fetch.bind(window);
window.fetch = (url, init) => {
  const entry = { method: (init && init.method) || 'GET', url: String(url), start: performance.now(), end: 0, status: 0, overlap: mock.log.some((e) => !e.end) };
  mock.log.push(entry);
  const done = (status) => {
    entry.end = performance.now();
    entry.status = status;
  };
  return realFetch(url, init).then(
    (response) => {
      done(response.status);
      return response;
    },
    (err) => {
      done(-1);
      throw err;
    }
  );
};

initializeIcons();

const names = (query.get('wp') || '').split(',').filter(Boolean);
const root = document.getElementById('root');

for (const name of names) {
  const host = document.createElement('div');
  host.setAttribute('data-wp', name);
  root.appendChild(host);

  if (name === 'AlertBanner') {
    const Customizer = require('../../lib/extensions/alertBanner/AlertBannerApplicationCustomizer').default;
    const handlers = [];
    const top = document.createElement('div');
    top.id = 'top-placeholder';
    document.body.insertBefore(top, document.body.firstChild);
    const instance = new Customizer();
    instance.context = {
      spHttpClient: mock.spHttpClient,
      pageContext,
      placeholderProvider: { changedEvent: { add: (owner, fn) => handlers.push(fn.bind(owner)) }, tryCreateContent: () => ({ domElement: top }) },
      application: { navigatedEvent: { add: (owner, fn) => handlers.push(fn.bind(owner)) } }
    };
    instance.onInit();
    // SharePoint raises both events right after start-up; do the same.
    handlers.forEach((fn) => fn());
    window.__navigate = () => handlers.forEach((fn) => fn());
    continue;
  }

  const Component = COMPONENTS[name]();
  const properties = Object.assign({ title: '' }, DEFAULT_PROPS[name] || {}, JSON.parse(query.get('props') || '{}'));
  ReactDom.render(
    React.createElement(Component, { context: { spHttpClient: mock.spHttpClient, pageContext, instanceId: 'instance-1' }, properties }),
    host
  );
}
