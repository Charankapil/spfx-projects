import { SPHttpClient } from '@microsoft/sp-http';

import { IFieldDefinition, ListHelper } from '../../../common/services/ListHelper';
import { SpRestClient } from '../../../common/services/SpRestClient';
import { IScorecardRun, ISnapshot } from '../models/IScorecard';
import { aggregate, scoreOf } from './metrics';

const FIELDS: IFieldDefinition[] = [
  { name: 'EcmRunDate', xml: '<Field Type="DateTime" Name="EcmRunDate" StaticName="EcmRunDate" DisplayName="Scanned" Format="DateTime" />' },
  { name: 'EcmScore', xml: '<Field Type="Number" Name="EcmScore" StaticName="EcmScore" DisplayName="Score" Decimals="0" />' },
  { name: 'EcmMetadataPct', xml: '<Field Type="Number" Name="EcmMetadataPct" StaticName="EcmMetadataPct" DisplayName="Metadata complete %" Decimals="1" />' },
  { name: 'EcmLabelledPct', xml: '<Field Type="Number" Name="EcmLabelledPct" StaticName="EcmLabelledPct" DisplayName="Retention labelled %" Decimals="1" />' },
  { name: 'EcmStalePct', xml: '<Field Type="Number" Name="EcmStalePct" StaticName="EcmStalePct" DisplayName="Stale %" Decimals="1" />' },
  { name: 'EcmDocuments', xml: '<Field Type="Number" Name="EcmDocuments" StaticName="EcmDocuments" DisplayName="Documents" Decimals="0" />' },
  { name: 'EcmSites', xml: '<Field Type="Number" Name="EcmSites" StaticName="EcmSites" DisplayName="Sites" Decimals="0" />' },
  {
    name: 'EcmPayload',
    xml: '<Field Type="Note" Name="EcmPayload" StaticName="EcmPayload" DisplayName="Results (JSON)" RichText="FALSE" NumLines="6" UnlimitedLengthInDocumentLibrary="TRUE" />'
  }
];

/** Multiple lines of text in a list hold up to 63,999 characters. */
const MAX_PAYLOAD = 63000;

interface ISnapshotItem {
  Id: number;
  EcmPayload: string | null;
  Author?: { Title?: string };
}

/**
 * Scorecard snapshots live in a list on the Corporate ECM Hub: one item per
 * scan, with the headline numbers as columns (readable in the list itself)
 * and every site's counts as JSON, so trends can be recomputed for any hub.
 */
export class SnapshotStore {
  private rest: SpRestClient;
  private list: ListHelper;

  constructor(httpClient: SPHttpClient, siteUrl: string, private listTitle: string) {
    this.rest = new SpRestClient(httpClient, 'Governance Scorecard');
    this.list = new ListHelper(this.rest, siteUrl, listTitle);
  }

  /** Most recent first; empty when the list does not exist yet. */
  public async load(top = 24): Promise<ISnapshot[]> {
    const exists = await this.list.getList();
    if (!exists) {
      return [];
    }
    const json = await this.rest.getJson<{ value: ISnapshotItem[] }>(
      `${this.list.listApi}/items?$select=Id,EcmPayload,Author/Title&$expand=Author&$orderby=EcmRunDate desc&$top=${top}`
    );
    const snapshots: ISnapshot[] = [];
    (json.value || []).forEach((item) => {
      try {
        const run = JSON.parse(item.EcmPayload || '') as IScorecardRun;
        if (run && run.version === 1 && Array.isArray(run.sites)) {
          snapshots.push({ ...run, id: item.Id, savedBy: item.Author?.Title });
        }
      } catch {
        // A hand-edited or truncated item is skipped rather than breaking the scorecard.
      }
    });
    return snapshots;
  }

  public async save(run: IScorecardRun): Promise<void> {
    const payload = JSON.stringify(run);
    if (payload.length > MAX_PAYLOAD) {
      throw new Error(`These results are too large to save (${payload.length} characters). Reduce the required fields or the number of sites.`);
    }
    await this.list.ensure('Governance Scorecard snapshots. Written by the Governance Scorecard web part.', FIELDS, () => undefined);
    const total = aggregate(run.sites, run.requiredFields);
    const parts = scoreOf(total);
    const round = (v: number | undefined, digits: number): number | null => (v === undefined ? null : parseFloat(v.toFixed(digits)));
    // Typed JSON rather than display strings: numbers and dates then do not depend on the user's regional settings.
    await this.rest.postJson(`${this.list.listApi}/items`, {
      Title: `Scan ${run.runAt.substring(0, 16).replace('T', ' ')} UTC`,
      EcmRunDate: run.runAt,
      EcmScore: round(parts.score, 0),
      EcmMetadataPct: round(parts.metadata, 1),
      EcmLabelledPct: round(parts.labelled, 1),
      EcmStalePct: round(parts.stale, 1),
      EcmDocuments: total.documents,
      EcmSites: total.scannedSites,
      EcmPayload: payload
    });
  }
}
