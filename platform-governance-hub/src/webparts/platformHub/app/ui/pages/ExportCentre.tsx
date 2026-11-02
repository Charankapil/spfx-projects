import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Text } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Card, PageGrid } from '../components/Layout';
import { PageFrame } from '../components/PageFrame';
import { buildHealthSection, PAGES } from '../sections';
import { loadPrices } from '../exportHelpers';
import { ExportFormat, runExport } from '../components/ExportMenu';
import { IExportSection } from '../../export/tabular';
import { describeError } from '../../http/ApiClient';

const HEALTH = 'health';

export const ExportCentrePage: React.FC = () => {
  const { provider, envFilter, notify, theme } = useApp();
  const exportable = PAGES.filter((p) => !!p.build);
  const [selected, setSelected] = React.useState<Record<string, boolean>>(() => {
    const s: Record<string, boolean> = {};
    exportable.forEach((p) => (s[p.id] = p.group === 'Power Platform' || p.id === 'overview'));
    return s;
  });
  const [format, setFormat] = React.useState<ExportFormat>('xlsx');
  const [applyFilter, setApplyFilter] = React.useState(true);
  const [progress, setProgress] = React.useState<{ done: number; total: number; label: string } | undefined>(undefined);
  const [failures, setFailures] = React.useState<string[]>([]);

  const ids = Object.keys(selected).filter((k) => selected[k]);
  const generate = async (): Promise<void> => {
    const chosen = exportable.filter((p) => selected[p.id]);
    const total = chosen.length + (selected[HEALTH] ? 1 : 0);
    const sections: IExportSection[] = [];
    const failed: string[] = [];
    const prices = loadPrices();
    const filter = applyFilter ? envFilter : [];
    setFailures([]);
    let done = 0;
    for (const p of chosen) {
      setProgress({ done, total, label: `Collecting ${p.title}` });
      try {
        if (p.build) sections.push(await p.build(provider, filter, prices));
      } catch (e) {
        failed.push(`${p.title}: ${describeError(e)}`);
      }
      done++;
    }
    if (selected[HEALTH]) {
      setProgress({ done, total, label: 'Collecting service health' });
      try {
        sections.push(await buildHealthSection(provider));
      } catch (e) {
        failed.push(`Service health: ${describeError(e)}`);
      }
    }
    setProgress({ done: total, total, label: 'Writing file' });
    try {
      if (sections.length) {
        await runExport(format, 'Platform Governance Report', sections, provider.tenantLabel);
        notify(`Report exported with ${sections.length} section(s).`, failed.length ? 'info' : 'success');
      }
    } catch (e) {
      failed.push(`Writing the file: ${describeError(e)}`);
    } finally {
      setProgress(undefined);
      setFailures(failed);
    }
  };

  const groups = ['Overview', 'Power Platform', 'Microsoft 365'];
  return (
    <PageFrame id="export">
      <PageGrid>
        <Card title="1 · Pages" span={7}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: 16 }}>
            {groups.map((g) => (
              <div key={g} style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <Text variant="small" style={{ fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em', color: theme.semanticColors.bodySubtext }}>{g}</Text>
                {exportable.filter((p) => p.group === g).map((p) => (
                  <Checkbox key={p.id} label={p.title} checked={!!selected[p.id]} onChange={(_, v) => setSelected({ ...selected, [p.id]: !!v })} />
                ))}
                {g === 'Microsoft 365' && <Checkbox label="Service health & message center" checked={!!selected[HEALTH]} onChange={(_, v) => setSelected({ ...selected, [HEALTH]: !!v })} />}
              </div>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <DefaultButton text="Select all" onClick={() => { const s: Record<string, boolean> = { [HEALTH]: true }; exportable.forEach((p) => (s[p.id] = true)); setSelected(s); }} />
            <DefaultButton text="Clear" onClick={() => setSelected({})} />
          </div>
        </Card>
        <Card title="2 · Format and scope" span={5}>
          <ChoiceGroup
            selectedKey={format}
            onChange={(_, o) => o && setFormat(o.key as ExportFormat)}
            options={[
              { key: 'xlsx', text: 'Excel workbook: a summary sheet plus one sheet per table' },
              { key: 'pdf', text: 'PDF report pack: headline numbers and tables (first 500 rows each)' },
              { key: 'csv', text: 'CSV: one file per table' }
            ]}
          />
          <Checkbox
            label={envFilter.length ? `Only the ${envFilter.length} environment(s) selected in the toolbar` : 'Apply the toolbar environment filter (none selected: all environments)'}
            checked={applyFilter}
            onChange={(_, v) => setApplyFilter(!!v)}
          />
          <PrimaryButton text={progress ? 'Generating…' : 'Generate export'} iconProps={{ iconName: 'Download' }} disabled={!ids.length || !!progress} onClick={() => { generate().catch(() => undefined); }} />
          {progress && <ProgressIndicator label={progress.label} percentComplete={progress.total ? progress.done / progress.total : 0} />}
          {failures.length > 0 && (
            <MessageBar messageBarType={MessageBarType.warning} isMultiline>
              <strong>Some sections were left out:</strong>
              {failures.map((f, i) => (
                <div key={i}>{f}</div>
              ))}
            </MessageBar>
          )}
          <Text variant="small" style={{ color: theme.semanticColors.bodySubtext }}>
            For charts in a PDF, use the Export button on the page itself, which includes the charts shown on screen. Downloads stay on this device; nothing is sent to another service.
          </Text>
        </Card>
      </PageGrid>
    </PageFrame>
  );
};
