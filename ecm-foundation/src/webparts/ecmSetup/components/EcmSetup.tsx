import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { TextField } from '@fluentui/react/lib/TextField';

import { DEFAULT_BRAND, HUB_COLORS, IEcmBrand, isHexColor, resolveBrand } from '../../../common/theme/ecmTheme';
import { ILinkCheck, SetupService } from '../services/SetupService';
import styles from './EcmSetup.module.scss';

export interface IEcmSetupProps {
  service: SetupService;
  tenantId: string;
  siteTitle: string;
  isSiteOwner: boolean;
}

type StepId = 'list' | 'read' | 'tenantTheme' | 'applyTheme' | 'navigator' | 'links';
type Status = 'idle' | 'running' | 'done' | 'failed';

interface IStepState {
  status: Status;
  log: string[];
}

const STATUS_LABEL: { [s in Status]: string } = { idle: 'Not run', running: 'Running…', done: 'Done', failed: 'Needs attention' };
const LINK_LABEL: { [s in ILinkCheck['status']]: string } = { ok: 'OK', 'no-access': 'No access', missing: 'Not found', error: 'Check' };

const Step = (props: {
  number: number;
  title: string;
  who: string;
  state: IStepState;
  onRun: () => void;
  runLabel?: string;
  disabled?: boolean;
  children?: React.ReactNode;
}): JSX.Element => (
  <section className={styles.step}>
    <div className={styles.stepHead}>
      <span className={`${styles.stepNumber} ${props.state.status === 'done' ? styles.numberDone : ''}`} aria-hidden="true">
        {props.state.status === 'done' ? '✓' : props.number}
      </span>
      <div className={styles.stepTitle}>
        <h3>{props.title}</h3>
        <span className={styles.who}>{props.who}</span>
      </div>
      <span className={`${styles.status} ${styles[props.state.status]}`}>{STATUS_LABEL[props.state.status]}</span>
      <DefaultButton
        text={props.runLabel || 'Run'}
        disabled={props.disabled || props.state.status === 'running'}
        onClick={props.onRun}
      />
    </div>
    {props.children && <div className={styles.stepBody}>{props.children}</div>}
    {props.state.log.length > 0 && (
      <ul className={styles.log} aria-live="polite">
        {props.state.log.map((line, i) => (
          <li key={i}>{line}</li>
        ))}
      </ul>
    )}
  </section>
);

export const EcmSetup = (props: IEcmSetupProps): JSX.Element => {
  const { service } = props;
  const idle = (): IStepState => ({ status: 'idle', log: [] });
  const [steps, setSteps] = React.useState<{ [id in StepId]: IStepState }>({
    list: idle(),
    read: idle(),
    tenantTheme: idle(),
    applyTheme: idle(),
    navigator: idle(),
    links: idle()
  });
  const [includeRead, setIncludeRead] = React.useState(true);
  const [themeSites, setThemeSites] = React.useState('');
  const [brandBadge, setBrandBadge] = React.useState('AT');
  const [homeTitle, setHomeTitle] = React.useState('Corporate ECM Hub');
  const [links, setLinks] = React.useState<ILinkCheck[]>([]);
  const [listUrl, setListUrl] = React.useState<string | undefined>(undefined);
  const [runningAll, setRunningAll] = React.useState(false);
  const [primary, setPrimary] = React.useState(DEFAULT_BRAND.primary);
  const [dark, setDark] = React.useState(DEFAULT_BRAND.dark);
  const [logoUrl, setLogoUrl] = React.useState('');
  const [logoState, setLogoState] = React.useState<IStepState>({ status: 'idle', log: [] });
  const [logoBroken, setLogoBroken] = React.useState(false);
  const primaryValid = isHexColor(primary);
  const darkValid = isHexColor(dark);
  const brand: IEcmBrand = resolveBrand({ primary, dark, logoUrl });

  // Pick up the brand saved by an earlier run, when this person can read it.
  React.useEffect(() => {
    service
      .loadSavedSettings()
      .then((saved) => {
        if (saved.brandPrimary && isHexColor(saved.brandPrimary)) {
          setPrimary(saved.brandPrimary);
        }
        if (saved.brandDark && isHexColor(saved.brandDark)) {
          setDark(saved.brandDark);
        }
        if (saved.logoUrl) {
          setLogoUrl(saved.logoUrl);
        }
        if (saved.brandBadge) {
          setBrandBadge(saved.brandBadge);
        }
        if (saved.homeTitle) {
          setHomeTitle(saved.homeTitle);
        }
      })
      .catch(() => undefined);
  }, [service]);

  const onLogoChosen = async (file: File | undefined): Promise<void> => {
    if (!file) {
      return;
    }
    setLogoState({ status: 'running', log: [] });
    const lines: string[] = [];
    try {
      const url = await service.uploadLogo(file, includeRead, props.tenantId, (l) => lines.push(l));
      setLogoUrl(url);
      setLogoBroken(false);
      lines.push('Run step 5 to show it in the navigator on every site.');
      setLogoState({ status: 'done', log: lines });
    } catch (err) {
      setLogoState({ status: 'failed', log: lines.concat([(err as Error).message]) });
    }
  };

  const loadThemeSites = React.useCallback(async () => {
    setThemeSites((await service.getThemeTargets()).join('\n'));
    setListUrl(await service.getNavigationListUrl());
  }, [service]);

  React.useEffect(() => {
    loadThemeSites().catch(() => undefined);
  }, [loadThemeSites]);

  const update = (id: StepId, change: (s: IStepState) => IStepState): void =>
    setSteps((all) => ({ ...all, [id]: change(all[id]) }));

  /** Runs one step with its own log; returns whether it succeeded. */
  const run = async (id: StepId, work: (log: (line: string) => void) => Promise<unknown>): Promise<boolean> => {
    update(id, () => ({ status: 'running', log: [] }));
    const log = (line: string): void => update(id, (s) => ({ ...s, log: s.log.concat([line]) }));
    try {
      await work(log);
      update(id, (s) => ({ ...s, status: 'done' }));
      return true;
    } catch (err) {
      log((err as Error).message);
      update(id, (s) => ({ ...s, status: 'failed' }));
      return false;
    }
  };

  const runList = (): Promise<boolean> =>
    run('list', async (log) => {
      await service.ensureNavigationList(log);
      await loadThemeSites();
    });
  const runRead = (): Promise<boolean> => run('read', (log) => service.grantEveryoneRead(props.tenantId, log));
  const runTenantTheme = (): Promise<boolean> => run('tenantTheme', (log) => service.registerTenantTheme(brand, log));
  const runApplyTheme = (sites?: string): Promise<boolean> =>
    run('applyTheme', async (log) => {
      const result = await service.applyTheme(brand, (sites ?? themeSites).split(/\r?\n/), log);
      log(`${result.applied} site(s) themed${result.failed ? `, ${result.failed} skipped` : ''}.`);
      if (!result.applied) {
        throw new Error('No site could be themed.');
      }
    });
  const runNavigator = (): Promise<boolean> => run('navigator', (log) => service.configureNavigator(brand, brandBadge, homeTitle, log));
  const runLinks = (): Promise<boolean> =>
    run('links', async (log) => {
      const found: ILinkCheck[] = [];
      setLinks([]);
      await service.checkLinks((r) => {
        found.push(r);
        setLinks(found.slice());
      });
      const bad = found.filter((r) => r.status === 'missing' || r.status === 'error').length;
      log(`${found.length} sites checked: ${found.length - bad} reachable, ${bad} to correct in the navigation list.`);
    });

  const runAll = async (): Promise<void> => {
    setRunningAll(true);
    try {
      if (!(await runList())) {
        return;
      }
      if (includeRead) {
        await runRead();
      }
      await runTenantTheme();
      // Re-read the targets: the list may have just been created.
      const sites = (await service.getThemeTargets()).join('\n');
      setThemeSites(sites);
      await runApplyTheme(sites);
      await runNavigator();
      await runLinks();
    } finally {
      setRunningAll(false);
    }
  };

  const badLinks = links.filter((l) => l.status !== 'ok');

  return (
    <div className={styles.setup}>
      <header className={styles.header}>
        <div>
          <h2>ECM Setup</h2>
          <p>
            Sets up the ECM Hub Navigator and the ECM theme with <b>{props.siteTitle}</b> as the Corporate ECM Hub. Everything runs as
            you, through SharePoint itself: no app registration or client id. Every step is safe to run again.
          </p>
        </div>
        <PrimaryButton text={runningAll ? 'Running…' : 'Run all steps'} disabled={runningAll} onClick={() => runAll()} />
      </header>

      {!props.isSiteOwner && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          You are not an owner of this site, so the first steps will fail. Ask a site owner to open this page.
        </MessageBar>
      )}

      <section className={styles.brandCard}>
        <div className={styles.brandHead}>
          <div>
            <h3>Brand</h3>
            <p>
              Used by steps 3 to 5. Enter the exact colours from the American Tower brand guide and upload the official logo from your
              brand team or media kit. The navigator and the SharePoint theme are built from these.
            </p>
          </div>
        </div>
        <div className={styles.brandGrid}>
          <div className={styles.brandFields}>
            <div>
              <span className={styles.fieldLabel}>Logo</span>
              <div className={styles.logoRow}>
                <span className={styles.logoPreview}>
                  {brand.logoUrl && !logoBroken ? (
                    <img src={brand.logoUrl} alt="Current logo" onError={() => setLogoBroken(true)} />
                  ) : (
                    <span>{brand.logoUrl ? 'Logo could not be loaded' : 'No logo yet'}</span>
                  )}
                </span>
                <label className={styles.fileButton}>
                  Upload logo…
                  <input
                    type="file"
                    accept="image/png,image/svg+xml,image/jpeg,image/gif,image/webp"
                    onChange={(e) => {
                      const file = e.target.files ? e.target.files[0] : undefined;
                      // Clear it so choosing the same file again still uploads.
                      e.target.value = '';
                      onLogoChosen(file).catch(() => undefined);
                    }}
                  />
                </label>
              </div>
              <span className={styles.hint}>SVG or transparent PNG, designed for a white background. Shown 32 px high.</span>
              {logoState.log.length > 0 && (
                <ul className={`${styles.miniLog} ${logoState.status === 'failed' ? styles.miniLogError : ''}`}>
                  {logoState.log.map((l, i) => (
                    <li key={i}>{l}</li>
                  ))}
                </ul>
              )}
            </div>
            <div className={styles.colorRow}>
              <div className={styles.colorField}>
                <label htmlFor="ecm-primary" className={styles.fieldLabel}>
                  Brand colour
                </label>
                <div className={styles.colorInputs}>
                  <input id="ecm-primary-swatch" type="color" aria-label="Pick brand colour" value={brand.primary} onChange={(e) => setPrimary(e.target.value.toUpperCase())} />
                  <input id="ecm-primary" className={styles.hexInput} value={primary} onChange={(e) => setPrimary(e.target.value)} maxLength={7} />
                </div>
                <span className={primaryValid ? styles.hint : styles.hintError}>
                  {primaryValid ? 'Buttons, links and the accent line.' : `Use a 6-digit hex code like #C91F3C; ${DEFAULT_BRAND.primary} is used until then.`}
                </span>
              </div>
              <div className={styles.colorField}>
                <label htmlFor="ecm-dark" className={styles.fieldLabel}>
                  Dark colour
                </label>
                <div className={styles.colorInputs}>
                  <input id="ecm-dark-swatch" type="color" aria-label="Pick dark colour" value={brand.dark} onChange={(e) => setDark(e.target.value.toUpperCase())} />
                  <input id="ecm-dark" className={styles.hexInput} value={dark} onChange={(e) => setDark(e.target.value)} maxLength={7} />
                </div>
                <span className={darkValid ? styles.hint : styles.hintError}>
                  {darkValid ? 'The hub band behind the tabs.' : `Use a 6-digit hex code like #1E2227; ${DEFAULT_BRAND.dark} is used until then.`}
                </span>
              </div>
            </div>
          </div>

          <div className={styles.preview} aria-label="Navigator preview" role="img">
            <div className={styles.previewAccent} style={{ background: brand.primary }} />
            <div className={styles.previewTop}>
              {brand.logoUrl && !logoBroken ? (
                <img src={brand.logoUrl} alt="" />
              ) : (
                <span className={styles.previewBadge} style={{ background: brand.primary }}>
                  {brandBadge || 'AT'}
                </span>
              )}
              <span className={styles.previewDivider} />
              <span className={styles.previewTitle}>{homeTitle}</span>
              <span className={styles.previewSpacer} />
              <span className={styles.previewButton} style={{ background: brand.primary }}>
                Request a site
              </span>
            </div>
            <div className={styles.previewBand} style={{ background: brand.dark }}>
              {Object.keys(HUB_COLORS)
                .slice(0, 4)
                .map((hub, i) => (
                  <span key={hub} className={`${styles.previewTab} ${i === 0 ? styles.previewTabOn : ''}`}>
                    <i style={{ background: HUB_COLORS[hub] }} />
                    {hub}
                  </span>
                ))}
            </div>
          </div>
        </div>
      </section>

      <Step
        number={1}
        title="Create the navigation list"
        who="Site owner of this site"
        state={steps.list}
        onRun={() => runList()}
      >
        Creates the <b>ECM Navigation</b> list with its columns and fills it with the six functional hubs, 21 department sites, hub
        resources and shared services from the ECM strategy. Existing entries are left as they are.
        {listUrl && (
          <>
            {' '}
            <a href={listUrl} target="_blank" rel="noopener noreferrer" data-interception="off">
              Open the list
            </a>
          </>
        )}
      </Step>

      <Step
        number={2}
        title="Let everyone read the menu"
        who="Site owner of this site"
        state={steps.read}
        onRun={() => runRead()}
      >
        The navigator can only show what the person viewing a page can read. This gives <b>Everyone except external users</b> read
        access to the navigation list only (not the site). Skip it if the hub&apos;s Visitors group already contains every employee.
        <Checkbox
          className={styles.inline}
          label="Include this step in “Run all steps”"
          checked={includeRead}
          onChange={(_, v) => setIncludeRead(!!v)}
        />
      </Step>

      <Step
        number={3}
        title="Register the ECM theme"
        who="SharePoint administrator"
        state={steps.tenantTheme}
        onRun={() => runTenantTheme()}
      >
        Adds <b>ECM Theme</b> to “Change the look” on every site, so owners of new sites can pick it. Not required: step 4 works without
        it.
      </Step>

      <Step
        number={4}
        title="Apply the ECM theme"
        who="Owner of each site listed"
        state={steps.applyTheme}
        onRun={() => runApplyTheme()}
      >
        Colours SharePoint&apos;s own header, buttons and links with the ECM palette on these sites. Sites you don&apos;t own are skipped
        and reported.
        <TextField
          className={styles.inline}
          label="Sites (one URL per line)"
          multiline
          rows={5}
          value={themeSites}
          onChange={(_, v) => setThemeSites(v || '')}
        />
        <DefaultButton className={styles.inline} text="Reload from navigation list" onClick={() => loadThemeSites()} />
      </Step>

      <Step
        number={5}
        title="Point the navigator at this hub"
        who="SharePoint administrator (app catalog owner)"
        state={steps.navigator}
        onRun={() => runNavigator()}
      >
        Tells the ECM Hub Navigator on every site to read its menu from this site. Needs the package deployed with “Make this app
        available to all sites” first.
        <div className={styles.fieldRow}>
          <TextField label="Brand badge" value={brandBadge} onChange={(_, v) => setBrandBadge(v || '')} maxLength={3} />
          <TextField label="Hub title" value={homeTitle} onChange={(_, v) => setHomeTitle(v || '')} />
        </div>
      </Step>

      <Step number={6} title="Check the links" who="Anyone" state={steps.links} onRun={() => runLinks()}>
        Opens every site the menu points to. The seeded URLs follow <code>/sites/&lt;name&gt;</code>; correct any that are not found in
        the navigation list.
        {steps.links.status === 'running' && <Spinner className={styles.inline} size={SpinnerSize.small} label="Checking…" />}
        {badLinks.length > 0 && (
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Entry</th>
                  <th>Type</th>
                  <th>URL</th>
                  <th>Result</th>
                </tr>
              </thead>
              <tbody>
                {badLinks.map((l) => (
                  <tr key={l.type + l.title}>
                    <td>{l.title}</td>
                    <td>{l.type}</td>
                    <td className={styles.url}>{l.url}</td>
                    <td>
                      <span className={`${styles.linkStatus} ${styles[l.status.replace('-', '')]}`}>{LINK_LABEL[l.status]}</span> {l.detail}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Step>
    </div>
  );
};
