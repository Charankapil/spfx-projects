/**
 * Generates the repetitive per-web-part files (web part class, manifest) and
 * config/config.json from one table, so the web parts cannot drift
 * apart. Run with: node tools/generate-webparts.js
 */
'use strict';

const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');

// kind: 'text' | 'multiline' | 'dropdown' | 'toggle'
const WEB_PARTS = [
  {
    dir: 'storageHogFinder', name: 'StorageHogFinder', id: '86f676bf-72a3-4484-b9f7-fbbb7db23db7',
    title: 'Storage Hog Finder', icon: 'Database',
    description: 'Shows which libraries, folders and version histories use the most storage.',
    props: []
  },
  {
    dir: 'readAcknowledge', name: 'ReadAcknowledge', id: '94a4b84f-1f9c-41d6-9733-1e6a310a202f',
    title: 'Read & Acknowledge', icon: 'ReadingMode',
    description: 'People confirm they have read each policy; owners see who has.',
    props: [{ key: 'libraryTitle', label: 'Policy library name', kind: 'text', default: 'Policies' }]
  },
  {
    dir: 'staleContentRadar', name: 'StaleContentRadar', id: '224affcd-4925-4cd0-85e1-c61ea0955e51',
    title: 'Stale Content Radar', icon: 'History',
    description: 'Finds files nobody has changed for a long time.',
    props: [{
      key: 'months', label: 'Stale after', kind: 'dropdown', default: '12',
      options: [['6', '6 months'], ['12', '1 year'], ['24', '2 years'], ['36', '3 years'], ['60', '5 years']]
    }]
  },
  {
    dir: 'recycleBinExplorer', name: 'RecycleBinExplorer', id: 'a00fd756-a456-4e5c-a987-1c66da5d2948',
    title: 'Recycle Bin Explorer', icon: 'RecycleBin',
    description: 'Search the recycle bin: what was deleted, when, and by whom.',
    props: []
  },
  {
    dir: 'bulkMetadataEditor', name: 'BulkMetadataEditor', id: '335a9572-7ad3-4ae6-94cc-454a369c5aa3',
    title: 'Bulk Metadata Editor', icon: 'Table',
    description: 'Edit the properties of many files at once in a grid.',
    props: []
  },
  {
    dir: 'duplicateFileDetector', name: 'DuplicateFileDetector', id: 'c33ef2fd-69db-4107-9f03-7207b787d1a8',
    title: 'Duplicate File Detector', icon: 'Copy',
    description: 'Finds files with the same name and size.',
    props: []
  },
  {
    dir: 'myFiles', name: 'MyFiles', id: '62d6e1d7-c52a-4909-bbba-236c546ae240',
    title: 'My Files', icon: 'Contact',
    description: 'Files you have checked out and files you worked on recently in this site.',
    props: [{
      key: 'rowCount', label: 'Files to show', kind: 'dropdown', default: '10',
      options: [['5', '5'], ['10', '10'], ['20', '20'], ['50', '50']]
    }]
  },
  {
    dir: 'pulsePoll', name: 'PulsePoll', id: 'faee9677-c848-41c5-8ab9-b52cff4c3a8a',
    title: 'Pulse Poll', icon: 'Poll',
    description: 'A one-question poll with live results.',
    props: [
      { key: 'question', label: 'Question', kind: 'text', default: 'How was your week?' },
      { key: 'options', label: 'Answers (one per line)', kind: 'multiline', default: 'Great\nGood\nOkay\nNot good' },
      { key: 'pollKey', label: 'Poll key (change it to start a fresh poll)', kind: 'text', default: '' }
    ]
  },
  {
    dir: 'templateGallery', name: 'TemplateGallery', id: 'c6d585f2-4a35-4aef-bf71-a4532004e379',
    title: 'Template Gallery', icon: 'DocumentSet',
    description: 'Start a new document from an approved template.',
    props: [{ key: 'libraryTitle', label: 'Template library name', kind: 'text', default: 'Templates' }]
  },
  {
    dir: 'orgChart', name: 'OrgChart', id: 'aba0f5f3-efce-41aa-bf78-00545dca8d41',
    title: 'Org Chart', icon: 'Org',
    description: 'Shows anyone\'s manager, team and peers, from Microsoft 365 profiles.',
    props: [
      {
        key: 'managerLevels', label: 'Manager levels shown above a person', kind: 'dropdown', default: '2',
        options: [['1', '1'], ['2', '2'], ['3', '3'], ['5', '5']]
      },
      { key: 'showPhotos', label: 'Show profile photos', kind: 'toggle', default: true }
    ]
  },
  {
    dir: 'faqAccordion', name: 'FaqAccordion', id: '4c1665d3-9ab6-4d0c-936d-bd079670964c',
    title: 'FAQ', icon: 'Unknown',
    description: 'Searchable questions and answers kept in a list.',
    props: []
  },
  {
    dir: 'whoHasAccess', name: 'WhoHasAccess', id: '4c13e94f-864a-45d7-a97c-4da888fb346b',
    title: 'Who Has Access?', icon: 'Permissions',
    description: 'Shows who can reach this site and what is shared outside it.',
    props: []
  },
  {
    dir: 'brokenLinkChecker', name: 'BrokenLinkChecker', id: '98e53560-1ef3-4c41-b8a5-d32d7bbb6d0b',
    title: 'Broken Link Checker', icon: 'RemoveLink',
    description: 'Checks the links on this site\'s pages and lists the dead ones.',
    props: [{
      key: 'maxLinks', label: 'Most links to check per run', kind: 'dropdown', default: '300',
      options: [['100', '100'], ['300', '300'], ['600', '600']]
    }]
  },
  {
    dir: 'alertBannerSettings', name: 'AlertBannerSettings', id: '6390b217-0232-467d-8894-1b3411aa2f2d',
    title: 'Alert Banner Settings', icon: 'Megaphone',
    description: 'Turns the site-wide alert banner on or off and manages its alerts.',
    props: []
  }
];

const EXTENSIONS = [
  { dir: 'alertBanner', name: 'AlertBannerApplicationCustomizer', id: 'b2be5f7c-564a-4ad3-ba1a-16ab1b12a70e' }
];

function write(file, content) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, content);
}

function paneField(prop) {
  if (prop.kind === 'toggle') {
    return `      PropertyPaneToggle('${prop.key}', { label: '${prop.label}' })`;
  }
  if (prop.kind === 'dropdown') {
    const options = prop.options.map(([key, text]) => `{ key: '${key}', text: '${text}' }`).join(', ');
    return `      PropertyPaneDropdown('${prop.key}', {\n        label: '${prop.label}',\n        options: [${options}]\n      })`;
  }
  return `      PropertyPaneTextField('${prop.key}', { label: '${prop.label}'${prop.kind === 'multiline' ? ', multiline: true, rows: 6' : ''} })`;
}

for (const wp of WEB_PARTS) {
  const dir = path.join(root, 'src', 'webparts', wp.dir);
  const usesDropdown = wp.props.some((p) => p.kind === 'dropdown');
  const usesToggle = wp.props.some((p) => p.kind === 'toggle');
  const imports = ['IPropertyPaneField', usesDropdown ? 'PropertyPaneDropdown' : undefined, 'PropertyPaneTextField', usesToggle ? 'PropertyPaneToggle' : undefined]
    .filter(Boolean)
    .join(', ');
  const fields = ["      PropertyPaneTextField('title', { label: strings.TitleFieldLabel })"]
    .concat(wp.props.map(paneField))
    .join(',\n');

  write(
    path.join(dir, `${wp.name}WebPart.ts`),
    `// Generated by tools/generate-webparts.js - edit the table there, not this file.
import * as React from 'react';
import { ${imports} } from '@microsoft/sp-property-pane';

import * as strings from 'SiteToolkitStrings';
import { IToolkitComponentProps, ToolkitWebPart } from '../../common/ToolkitWebPart';
import { I${wp.name}Props, ${wp.name} } from './${wp.name}';

export default class ${wp.name}WebPart extends ToolkitWebPart<I${wp.name}Props> {
  protected readonly component: React.ComponentType<IToolkitComponentProps<I${wp.name}Props>> = ${wp.name};

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected get fields(): IPropertyPaneField<any>[] {
    return [
${fields}
    ];
  }
}
`
  );

  const properties = { title: wp.title };
  for (const prop of wp.props) {
    properties[prop.key] = prop.default;
  }
  write(
    path.join(dir, `${wp.name}WebPart.manifest.json`),
    JSON.stringify(
      {
        $schema: 'https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json',
        id: wp.id,
        alias: `${wp.name}WebPart`,
        componentType: 'WebPart',
        version: '*',
        manifestVersion: 2,
        requiresCustomScript: false,
        supportedHosts: ['SharePointWebPart', 'SharePointFullPage'],
        preconfiguredEntries: [
          {
            groupId: '5c03119e-3074-46fd-976b-c60198311f70',
            group: { default: 'Site Toolkit' },
            title: { default: wp.title },
            description: { default: wp.description },
            officeFabricIconFontName: wp.icon,
            properties
          }
        ]
      },
      null,
      2
    ) + '\n'
  );
}

const kebab = (name) => name.replace(/([a-z0-9])([A-Z])/g, '$1-$2').toLowerCase();
const bundles = {};
for (const wp of WEB_PARTS) {
  bundles[`${kebab(wp.name)}-web-part`] = {
    components: [
      {
        entrypoint: `./lib/webparts/${wp.dir}/${wp.name}WebPart.js`,
        manifest: `./src/webparts/${wp.dir}/${wp.name}WebPart.manifest.json`
      }
    ]
  };
}
for (const ext of EXTENSIONS) {
  bundles[kebab(ext.name)] = {
    components: [
      {
        entrypoint: `./lib/extensions/${ext.dir}/${ext.name}.js`,
        manifest: `./src/extensions/${ext.dir}/${ext.name}.manifest.json`
      }
    ]
  };
}

write(
  path.join(root, 'config', 'config.json'),
  JSON.stringify(
    {
      $schema: 'https://developer.microsoft.com/json-schemas/core-build/config.2.0.schema.json',
      version: '2.0',
      bundles,
      externals: {},
      localizedResources: { SiteToolkitStrings: 'lib/loc/{locale}.js' }
    },
    null,
    2
  ) + '\n'
);

console.log(`Generated ${WEB_PARTS.length} web parts and ${EXTENSIONS.length} extension bundle(s).`);
