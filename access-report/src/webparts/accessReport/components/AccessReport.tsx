import * as React from 'react';
import { MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import { SPClient } from '../core/SPClient';
import { AccessScanner } from '../services/AccessScan';
import { PeoplePicker } from '../services/PeoplePicker';
import { SiteLister } from '../services/Sites';
import { ISettings, IStoredSubject, Store } from '../services/Store';
import styles from './AccessReport.module.scss';
import { ReportsTab } from './ReportsTab';
import { RunTab } from './RunTab';
import { ScheduleTab } from './ScheduleTab';
import { SetupTab } from './SetupTab';
import { Card, errorText } from './ui';

export interface IAccessReportProps {
  client: SPClient;
  homeWebUrl: string;
  heading: string;
  /** Entra tenant id, for the "Everyone except external users" claim; empty when unknown. */
  tenantId: string;
  version: string;
}

export interface IServices {
  store: Store;
  scanner: AccessScanner;
  lister: SiteLister;
  picker: PeoplePicker;
  homeWebUrl: string;
  tenantId: string;
}

export interface IConfig {
  settings: ISettings;
  subjects: IStoredSubject[];
}

/**
 * Access Report: who can reach what, as a file in a SharePoint folder - on demand from this
 * page, and every month from the Power Automate flow that reads the same configuration.
 */
export const AccessReport: React.FC<IAccessReportProps> = (props) => {
  const services = React.useMemo<IServices>(
    () => ({
      store: new Store(props.client, props.homeWebUrl),
      scanner: new AccessScanner(props.client, props.homeWebUrl),
      lister: new SiteLister(props.client),
      picker: new PeoplePicker(props.client, props.homeWebUrl),
      homeWebUrl: props.homeWebUrl,
      tenantId: props.tenantId
    }),
    [props.client, props.homeWebUrl, props.tenantId]
  );
  const [missing, setMissing] = React.useState<string[] | undefined>();
  const [config, setConfig] = React.useState<IConfig | undefined>();
  const [busy, setBusy] = React.useState<string | undefined>('Loading…');
  const [error, setError] = React.useState<string | undefined>();
  const [tab, setTab] = React.useState('run');
  const alive = React.useRef(true);
  React.useEffect(
    () => () => {
      alive.current = false;
    },
    []
  );

  const load = React.useCallback(async (): Promise<void> => {
    setError(undefined);
    try {
      const m = await services.store.missing();
      if (!alive.current) {
        return;
      }
      setMissing(m);
      if (!m.length) {
        const [settings, subjects] = [await services.store.settings(), await services.store.subjects()];
        if (alive.current) {
          setConfig({ settings, subjects });
        }
      }
    } catch (e) {
      if (alive.current) {
        setError(errorText(e));
      }
    } finally {
      if (alive.current) {
        setBusy(undefined);
      }
    }
  }, [services]);

  React.useEffect(() => {
    load().catch(() => undefined);
  }, [load]);

  const setUp = async (): Promise<void> => {
    setBusy('Setting up…');
    setError(undefined);
    try {
      await services.store.provision((text) => alive.current && setBusy(text));
      await load();
    } catch (e) {
      if (alive.current) {
        setError(`Set-up did not finish: ${errorText(e)} You need to be an owner of this site to create the lists.`);
        setBusy(undefined);
      }
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <h2>{props.heading}</h2>
        <span className={`${styles.muted} ${styles.small}`}>
          Runs as you, with SharePoint permissions only · v{props.version}
        </span>
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}

      {busy && !config ? (
        <div className={styles.tab}>
          <Spinner size={SpinnerSize.medium} label={busy} labelPosition="right" style={{ justifyContent: 'flex-start' }} />
        </div>
      ) : missing && missing.length > 0 ? (
        <div className={styles.tab}>
          <Card title="Set up this site for the Access Report">
            <p className={styles.hint}>
              The report keeps its configuration and its files in this site: two small lists (who to report on, and the options) and a document library for the report files. They are created
              once, with your own permissions.
            </p>
            <p className={styles.hint}>Still to create: {missing.join(', ')}.</p>
            <PrimaryButton onClick={setUp} disabled={!!busy}>
              {busy || 'Set up'}
            </PrimaryButton>
          </Card>
        </div>
      ) : config ? (
        <Pivot selectedKey={tab} onLinkClick={(i) => i && setTab(String(i.props.itemKey))} overflowBehavior="menu">
          <PivotItem headerText="Run report" itemKey="run" alwaysRender>
            {/* Kept mounted so a run in progress survives a look at the other tabs. */}
            <RunTab services={services} config={config} onOpenTab={setTab} />
          </PivotItem>
          <PivotItem headerText="Who and where" itemKey="setup">
            <SetupTab services={services} config={config} onChange={setConfig} />
          </PivotItem>
          <PivotItem headerText="Saved reports" itemKey="reports">
            <ReportsTab services={services} config={config} />
          </PivotItem>
          <PivotItem headerText="Monthly schedule" itemKey="schedule">
            <ScheduleTab services={services} config={config} onChange={setConfig} />
          </PivotItem>
        </Pivot>
      ) : null}
    </div>
  );
};
