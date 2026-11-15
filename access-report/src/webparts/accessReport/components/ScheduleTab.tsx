import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import { FLOW_NAME } from '../flow/definition';
import { LEGACY_FILE, legacyPackage, SOLUTION_FILE, solutionPackage } from '../flow/package';
import { downloadBlob } from '../services/reportCsv';
import styles from './AccessReport.module.scss';
import { IConfig, IServices } from './AccessReport';
import { Card, errorText, Facts } from './ui';

/** The monthly flow: download it with this site filled in, import it, and see how its last run went. */
export const ScheduleTab: React.FC<{ services: IServices; config: IConfig; onChange: (c: IConfig) => void }> = ({ services, config, onChange }) => {
  const [error, setError] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);
  const home = services.store.web;
  const lastRun = config.settings.lastRun;
  const active = config.subjects.filter((s) => s.active && s.login).length;

  const download = (name: string, build: (homeUrl: string) => Uint8Array): void => {
    try {
      downloadBlob(name, new Blob([build(home)], { type: 'application/zip' }));
    } catch (e) {
      setError(errorText(e));
    }
  };
  const refresh = async (): Promise<void> => {
    setBusy(true);
    setError(undefined);
    try {
      onChange({ settings: await services.store.settings(), subjects: config.subjects });
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.tab}>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}

      <Card
        title="Last scheduled run"
        right={
          <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={refresh} disabled={busy}>
            Refresh
          </DefaultButton>
        }
      >
        {lastRun ? (
          <>
            <p style={{ margin: '0 0 6px' }}>{lastRun}</p>
            {/^Started/.test(lastRun) && (
              <p className={`${styles.muted} ${styles.small}`} style={{ margin: 0 }}>
                “Started” with no “Finished” means the flow is still running, or stopped before it could save: its run history in Power Automate says which.
              </p>
            )}
          </>
        ) : (
          <p className={styles.muted} style={{ margin: 0 }}>
            The flow has not run yet. Once it is imported and switched on, it writes a line here each time it starts and finishes.
          </p>
        )}
      </Card>

      <Card title="Have it run on the first of every month">
        <p className={styles.hint}>
          “{FLOW_NAME}” is a Power Automate cloud flow that does what “Run report” does, without anyone opening this page. It uses only the standard SharePoint connector (“Send an HTTP request to
          SharePoint”), so it needs no premium licence, no app registration and no admin consent. It runs as the person whose SharePoint connection it uses and sees the sites that person can open.
        </p>
        <Facts
          items={[
            { label: 'Reads its configuration from', value: <span className={styles.code}>{home}</span> },
            { label: 'People and groups', value: `${active.toLocaleString()} included` },
            { label: 'Saves into', value: <span className={styles.code}>{config.settings.reportFolder}</span> }
          ]}
        />
        <div className={styles.actions} style={{ margin: '14px 0' }}>
          <PrimaryButton iconProps={{ iconName: 'Download' }} onClick={() => download(LEGACY_FILE, legacyPackage)}>
            Download the flow package
          </PrimaryButton>
          <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => download(SOLUTION_FILE, solutionPackage)}>
            Download as a solution
          </DefaultButton>
        </div>
        <ol className={styles.steps}>
          <li>
            Sign in to <strong>make.powerautomate.com</strong> as the account the report should run as: one that can open the sites to report on (a site collection administrator of them) and can
            edit this site.
          </li>
          <li>
            <strong>My flows → Import → Import Package (Legacy)</strong>, choose the downloaded <span className={styles.code}>{LEGACY_FILE}</span>, pick “Create as new” and select or create the
            SharePoint connection, then <strong>Import</strong>.
          </li>
          <li>
            Open the flow and switch it <strong>On</strong>. Press <strong>Run</strong> once to test it now; after that it runs at 03:00 UTC on the first of each month (change the trigger for another
            time).
          </li>
          <li>The report appears in the folder above and under “Saved reports”; this page shows when it started and finished.</li>
        </ol>
        <p className={`${styles.muted} ${styles.small}`}>
          Prefer solutions? “Download as a solution” gives the same flow for <strong>Solutions → Import solution</strong> (a Dataverse environment), with a connection reference to map during the
          import.
        </p>
      </Card>

      <Card title="Good to know">
        <div className={styles.note}>
          Nothing about the report is configured in the flow. It reads “Who and where” from this site each time it runs, so changes made here apply to the next run without touching Power Automate.
        </div>
        <ul className={styles.steps} style={{ marginTop: 12 }}>
          <li>
            <strong>How long it takes.</strong> The flow works through one site at a time, which is gentle on SharePoint: roughly two to three actions per person and site, three per group and site,
            and a few more per site when libraries are included. A few hundred sites finish within the hour; thousands of sites with many people can take several hours, which is fine for a flow
            (runs may last up to 30 days).
          </li>
          <li>
            <strong>Daily action limits.</strong> Every flow action counts against the owner’s Power Platform request allowance. If the tenant is large, narrow the run with the site filter, leave
            libraries out, or run the flow under an account with a Power Automate licence.
          </li>
          <li>
            <strong>Sites the account cannot open</strong> are not skipped silently: they are listed in a second file next to the report (ending in <span className={styles.code}>-unread.csv</span>).
          </li>
          <li>
            <strong>Moving the web part to another site?</strong> Download the package again from there: the site address is written into its first step (“Home site URL”), which can also be edited
            in the flow.
          </li>
        </ul>
      </Card>
    </div>
  );
};
