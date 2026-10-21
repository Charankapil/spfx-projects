// Local preview of Formwright against an in-memory SharePoint site.
// See preview/README.md.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import { DesignerOverlay } from '@fs/extensions/formwrightCommandSet/DesignerOverlay';
import { FormHost } from '@fs/common/components/renderer/FormHost';
import { fakeContext, fakeData, log } from './fakeSharePoint';
import sample from '../samples/purchase-request.forms.json';

initializeIcons('./fonts/');
window.__fake = { data: fakeData, log };
const params = new URLSearchParams(location.search);
const view = params.get('view') || 'designer';
const root = document.getElementById('root');
const listId = fakeData.LIST_IDS.requests;

// ?sample=1 starts with the purchase-request sample forms already saved.
if (params.get('sample')) {
  fakeData.files[`/sites/procurement/SiteAssets/Formwright/${listId}.json`] = JSON.stringify(sample);
}

if (view === 'designer') {
  // What "Customize forms" in the list's command bar opens.
  ReactDom.render(React.createElement(DesignerOverlay, { ctx: fakeContext, listId, listTitle: 'Purchase Requests', onClose: () => (document.title = 'closed') }), root);
} else {
  ReactDom.render(
    React.createElement(FormHost, {
      ctx: fakeContext,
      listId,
      mode: params.get('mode') || 'New',
      itemId: params.get('item') ? Number(params.get('item')) : undefined,
      onSaved: (id, _values, redirect, message) => {
        window.__saved = { id, redirect, message };
        document.title = `saved ${id}`;
        console.log('Saved item', id, message, window.__fake.log);
      },
      onCancel: () => (document.title = 'cancelled'),
      onEdit: () => (location.search = `?view=form&mode=Edit&item=${params.get('item')}${params.get('sample') ? '&sample=1' : ''}`)
    }),
    root
  );
}
