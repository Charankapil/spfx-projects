import { buildDefinition, CONNECTION_KEY, FLOW_NAME, SHAREPOINT_API } from './definition';
import { IZipEntry, zip } from './zip';

/**
 * The flow as something Power Automate can import, in its two formats:
 *
 *  - a package for My flows > Import > Import Package (Legacy): works in every environment,
 *    including the default one, and asks for the SharePoint connection during the import;
 *  - an unmanaged solution for Solutions > Import solution (needs a Dataverse environment):
 *    the flow arrives with a connection reference.
 *
 * Both carry the same definition; the ids are fixed so a later import updates the same flow.
 */

const FLOW_ID = '0d57c5f4-7be8-4a02-ab10-3dcbea515778';
const API_RESOURCE = 'bb9d3cea-0091-4f3a-a250-c70ed5a370da';
const CONNECTION_RESOURCE = '0f7f4f58-6d2c-4acb-b3ed-1879ef2c866e';
const TELEMETRY_ID = '5e0b7a52-6f0c-4c58-9a57-2f7d1f6c9e11';
const CONNECTION_REFERENCE = 'ar_sharedsharepointonline_accessreport';
const SOLUTION_NAME = 'AccessReportMonthly';
const VERSION = '1.0.0.0';
const ICON = 'https://connectoricons-prod.azureedge.net/releases/v1.0.1669/1.0.1669.3522/sharepointonline/icon.png';
const DESCRIPTION = 'On the first of every month: lists every site and library the chosen people and groups can reach and saves the report as a CSV in a SharePoint folder. Configured by the Access Report web part.';

export const LEGACY_FILE = 'AccessReport-monthly-flow.zip';
export const SOLUTION_FILE = 'AccessReportMonthly_1_0_0_0.zip';

const json = (value: unknown): string => JSON.stringify(value, null, 2);

export function legacyEntries(homeUrl: string): IZipEntry[] {
  const flowPath = `Microsoft.Flow/flows/${FLOW_ID}`;
  const manifest = {
    schema: '1.0',
    details: { displayName: FLOW_NAME, description: DESCRIPTION, createdTime: '2026-01-01T00:00:00.0000000Z', packageTelemetryId: TELEMETRY_ID, creator: 'N/A', sourceEnvironment: '' },
    resources: {
      [FLOW_ID]: {
        id: `/providers/Microsoft.Flow/flows/${FLOW_ID}`,
        name: FLOW_ID,
        type: 'Microsoft.Flow/flows',
        suggestedCreationType: 'New',
        creationType: 'Existing, New, Update',
        details: { displayName: FLOW_NAME },
        configurableBy: 'User',
        hierarchy: 'Root',
        dependsOn: [API_RESOURCE, CONNECTION_RESOURCE]
      },
      [API_RESOURCE]: {
        id: SHAREPOINT_API,
        name: CONNECTION_KEY,
        type: 'Microsoft.PowerApps/apis',
        suggestedCreationType: 'Existing',
        details: { displayName: 'SharePoint', iconUri: ICON },
        configurableBy: 'System',
        hierarchy: 'Child',
        dependsOn: []
      },
      [CONNECTION_RESOURCE]: {
        type: 'Microsoft.PowerApps/apis/connections',
        suggestedCreationType: 'Existing',
        creationType: 'Existing',
        details: { displayName: 'SharePoint connection', iconUri: ICON },
        configurableBy: 'User',
        hierarchy: 'Child',
        dependsOn: [API_RESOURCE]
      }
    }
  };
  const definition = {
    name: FLOW_ID,
    id: `/providers/Microsoft.Flow/flows/${FLOW_ID}`,
    type: 'Microsoft.Flow/flows',
    properties: {
      apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
      displayName: FLOW_NAME,
      definition: buildDefinition({ homeUrl }),
      connectionReferences: { [CONNECTION_KEY]: { connectionName: '', source: 'Embedded', id: SHAREPOINT_API, tier: 'NotSpecified' } },
      flowFailureAlertSubscribed: false,
      isManaged: false
    }
  };
  return [
    { name: 'manifest.json', text: json(manifest) },
    { name: 'Microsoft.Flow/flows/manifest.json', text: json({ packageSchemaVersion: '1.0', flowAssets: { assetPaths: [FLOW_ID] } }) },
    { name: `${flowPath}/definition.json`, text: json(definition) },
    { name: `${flowPath}/apisMap.json`, text: json({ [CONNECTION_KEY]: API_RESOURCE }) },
    { name: `${flowPath}/connectionsMap.json`, text: json({ [CONNECTION_KEY]: CONNECTION_RESOURCE }) }
  ];
}

const xmlText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function address(n: number): string {
  const nil = ['City', 'County', 'Country', 'Fax', 'FreightTermsCode', 'ImportSequenceNumber', 'Latitude', 'Line1', 'Line2', 'Line3', 'Longitude', 'Name', 'PostalCode', 'PostOfficeBox', 'PrimaryContactName'];
  const nilAfter = ['StateOrProvince', 'Telephone1', 'Telephone2', 'Telephone3', 'TimeZoneRuleVersionNumber', 'UPSZone', 'UTCOffset', 'UTCConversionTimeZoneCode'];
  const el = (name: string): string => `          <${name} xsi:nil="true"></${name}>`;
  return [
    '        <Address>',
    `          <AddressNumber>${n}</AddressNumber>`,
    '          <AddressTypeCode>1</AddressTypeCode>',
    ...nil.map(el),
    '          <ShippingMethodCode>1</ShippingMethodCode>',
    ...nilAfter.map(el),
    '        </Address>'
  ].join('\n');
}

export function solutionEntries(homeUrl: string): IZipEntry[] {
  const workflowFile = `Workflows/AccessReport-monthly-${FLOW_ID.toUpperCase()}.json`;
  const workflow = {
    properties: {
      connectionReferences: { [CONNECTION_KEY]: { runtimeSource: 'embedded', connection: { connectionReferenceLogicalName: CONNECTION_REFERENCE }, api: { name: CONNECTION_KEY } } },
      definition: buildDefinition({ homeUrl })
    },
    schemaVersion: '1.0.0.0'
  };
  const solution = `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.24014.198" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <SolutionManifest>
    <UniqueName>${SOLUTION_NAME}</UniqueName>
    <LocalizedNames>
      <LocalizedName description="${xmlText(FLOW_NAME)}" languagecode="1033" />
    </LocalizedNames>
    <Descriptions>
      <Description description="${xmlText(DESCRIPTION)}" languagecode="1033" />
    </Descriptions>
    <Version>${VERSION}</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>accessreport</UniqueName>
      <LocalizedNames>
        <LocalizedName description="Access Report" languagecode="1033" />
      </LocalizedNames>
      <Descriptions />
      <EMailAddress xsi:nil="true"></EMailAddress>
      <SupportingWebsiteUrl xsi:nil="true"></SupportingWebsiteUrl>
      <CustomizationPrefix>ar</CustomizationPrefix>
      <CustomizationOptionValuePrefix>48213</CustomizationOptionValuePrefix>
      <Addresses>
${address(1)}
${address(2)}
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type="29" id="{${FLOW_ID}}" behavior="0" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>`;
  const customizations = `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities></Entities>
  <Roles></Roles>
  <Workflows>
    <Workflow WorkflowId="{${FLOW_ID}}" Name="${xmlText(FLOW_NAME)}">
      <JsonFileName>/${workflowFile}</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>1</StateCode>
      <StatusCode>2</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>${VERSION}</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="${xmlText(FLOW_NAME)}" />
      </LocalizedNames>
    </Workflow>
  </Workflows>
  <FieldSecurityProfiles></FieldSecurityProfiles>
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <connectionreferences>
    <connectionreference connectionreferencelogicalname="${CONNECTION_REFERENCE}">
      <connectionreferencedisplayname>SharePoint (Access Report)</connectionreferencedisplayname>
      <connectorid>${SHAREPOINT_API}</connectorid>
      <iscustomizable>1</iscustomizable>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>
  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>`;
  const contentTypes =
    '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/octet-stream" /><Default Extension="json" ContentType="application/octet-stream" /></Types>';
  return [
    { name: '[Content_Types].xml', text: contentTypes },
    { name: 'solution.xml', text: solution },
    { name: 'customizations.xml', text: customizations },
    { name: workflowFile, text: json(workflow) }
  ];
}

export function legacyPackage(homeUrl: string): Uint8Array {
  return zip(legacyEntries(homeUrl));
}

export function solutionPackage(homeUrl: string): Uint8Array {
  return zip(solutionEntries(homeUrl));
}
