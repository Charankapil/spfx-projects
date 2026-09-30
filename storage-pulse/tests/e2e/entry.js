// Renders the compiled web part (lib/) against the SharePoint simulation.
// The scenario comes from window.__scenario, set by run.mjs before the page loads.
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import { StoragePulse } from '../../lib/webparts/storagePulse/components/StoragePulse.js';
import { createMockSharePoint } from '../mock/mockSharePoint.js';

initializeIcons();
const scenario = window.__scenario || {};
const sp = createMockSharePoint(scenario.mock || {});
window.__sp = sp;

function render(extra) {
  const props = {
    title: 'Storage Pulse',
    scope: 'siteCollection',
    thresholdMonths: 12,
    scanPermission: 'owners',
    scanAllowedPeople: [],
    scanMode: 'quick',
    // Tests run without real-world pacing; scenarios can override any of it.
    scanSpeed: 'fast',
    includeHidden: false,
    excludeSystemLibraries: false,
    excludedLibraries: [],
    staleAfterDays: 90,
    theme: { isDark: false },
    context: sp.context,
    ...(scenario.props || {}),
    ...(extra || {})
  };
  props.tuning = { autoRetryDelayMs: 0, profileOverride: { minGapMs: 0 }, ...((scenario.props && scenario.props.tuning) || {}), ...((extra && extra.tuning) || {}) };
  const root = document.getElementById('root');
  ReactDOM.unmountComponentAtNode(root);
  ReactDOM.render(React.createElement(StoragePulse, props), root);
}
window.__remount = (owner, extra) => {
  sp.context.pageContext.web.permissions.hasPermission = () => owner;
  sp.context.pageContext.legacyPageContext.isSiteAdmin = owner;
  render(extra);
};
render();
