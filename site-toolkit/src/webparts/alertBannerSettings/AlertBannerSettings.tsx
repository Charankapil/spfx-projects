import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, Spinner, SpinnerSize, TextField, Toggle } from '@fluentui/react';

import { ALERT_BANNER_COMPONENT_ID, ALERT_CACHE_MINUTES, ALERTS, alertStatus, clearAlertCache, IAlert, loadAlerts, SEVERITIES } from '../../common/alerts';
import { formatDate, fromDateInput, safeHref, toDateInput } from '../../common/format';
import { errorText } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { SetupCard } from '../../common/ui/SetupCard';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IAlertBannerSettingsProps {
  title: string;
}

interface ICustomAction {
  Id: string;
  ClientSideComponentId: string;
}

const STATUS_STYLE: Record<string, string> = {
  active: styles.pillGood,
  scheduled: styles.pillWarn,
  expired: ''
};

export const AlertBannerSettings: React.FC<IToolkitComponentProps<IAlertBannerSettingsProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const actionsApi = `${kit.sp.siteUrl}/_api/site/UserCustomActions`;

  const [alerts, setAlerts] = useState<IAlert[] | undefined>(undefined);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [actionId, setActionId] = useState<string | undefined>(undefined);
  const [bannerKnown, setBannerKnown] = useState(false);
  const [loading, setLoading] = useState(true);
  const [switching, setSwitching] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<string | undefined>(undefined);

  const [formOpen, setFormOpen] = useState(false);
  const [headline, setHeadline] = useState('');
  const [message, setMessage] = useState('');
  const [severity, setSeverity] = useState(SEVERITIES[0]);
  const [start, setStart] = useState(toDateInput(new Date()));
  const [end, setEnd] = useState('');
  const [link, setLink] = useState('');
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState<string | undefined>(undefined);

  const load = useCallback(() => {
    setLoading(true);
    setError(undefined);
    const run = async (): Promise<{ alerts: IAlert[] | undefined; actionId: string | undefined; known: boolean }> => {
      const list = await loadAlerts(kit.sp, true);
      let id: string | undefined;
      let known = true;
      try {
        const actions = await kit.sp.getJson<{ value: ICustomAction[] }>(`${actionsApi}?$select=Id,ClientSideComponentId`);
        const match = (actions.value || []).filter(
          (a) => (a.ClientSideComponentId || '').toLowerCase() === ALERT_BANNER_COMPONENT_ID
        )[0];
        id = match ? match.Id : undefined;
      } catch {
        known = false;
      }
      return { alerts: list, actionId: id, known };
    };
    run()
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        setAlerts(result.alerts);
        setNeedsSetup(result.alerts === undefined);
        setActionId(result.actionId);
        setBannerKnown(result.known);
        setLoading(false);
      })
      .catch((err) => {
        if (mounted.current) {
          setLoading(false);
          setError(`The alert settings could not be loaded: ${errorText(err)}`);
        }
      });
  }, [actionsApi, kit, mounted]);

  useEffect(() => {
    // People who cannot manage the site see a notice only; nothing is requested for them.
    if (kit.canManage) {
      load();
    }
  }, [kit.canManage, load]);

  const setup = useCallback(async () => {
    await kit.rootLists.ensure(ALERTS);
    clearAlertCache(kit.sp);
  }, [kit]);

  const switchBanner = useCallback(
    (on: boolean) => {
      if (switching) {
        return;
      }
      setSwitching(true);
      setError(undefined);
      const work: Promise<string | undefined> = on
        ? kit.sp
            .post<ICustomAction>(actionsApi, {
              Title: 'Site Toolkit Alert Banner',
              Name: 'SiteToolkitAlertBanner',
              Location: 'ClientSideExtension.ApplicationCustomizer',
              ClientSideComponentId: ALERT_BANNER_COMPONENT_ID,
              ClientSideComponentProperties: '{}'
            })
            .then((created) => (created ? created.Id : undefined))
        : actionId
        ? kit.sp.remove(`${actionsApi}('${actionId}')`).then(() => undefined)
        : Promise.resolve(undefined);
      work
        .then((id) => {
          if (!mounted.current) {
            return;
          }
          setSwitching(false);
          setActionId(on ? id || 'on' : undefined);
          setNotice(
            on
              ? 'The banner is on for this site collection. It appears after the next page load.'
              : 'The banner is off for this site collection.'
          );
          if (on && !id) {
            load();
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setSwitching(false);
            setError(
              `The banner could not be switched ${on ? 'on' : 'off'}: ${errorText(err)} Only a site collection administrator can do this.`
            );
          }
        });
    },
    [actionId, actionsApi, kit, load, mounted, switching]
  );

  const closeForm = useCallback(() => {
    setFormOpen(false);
    setHeadline('');
    setMessage('');
    setSeverity(SEVERITIES[0]);
    setStart(toDateInput(new Date()));
    setEnd('');
    setLink('');
    setFormError(undefined);
  }, []);

  const addAlert = useCallback(() => {
    if (saving) {
      return;
    }
    const startIso = fromDateInput(start);
    const endIso = end ? fromDateInput(end) : undefined;
    if (!headline.trim()) {
      setFormError('Enter a headline.');
      return;
    }
    if (!startIso || (end && !endIso)) {
      setFormError('Enter the dates as valid dates.');
      return;
    }
    if (endIso && endIso < startIso) {
      setFormError('The end date is before the start date.');
      return;
    }
    if (link.trim() && !safeHref(link)) {
      setFormError('The link must start with https:// or with / for a page on this tenant.');
      return;
    }
    const values: Record<string, unknown> = {
      Title: headline.trim().substring(0, 255),
      Message: message.trim(),
      Severity: severity,
      StartDate: startIso,
      LinkUrl: link.trim()
    };
    if (endIso) {
      // The alert stays up through the whole of its last day.
      const lastMoment = new Date(endIso);
      lastMoment.setHours(23, 59, 59, 0);
      values.EndDate = lastMoment.toISOString();
    }
    setSaving(true);
    setFormError(undefined);
    kit.rootLists
      .addItem(ALERTS.urlName, values)
      .then(() => {
        if (!mounted.current) {
          return;
        }
        setSaving(false);
        closeForm();
        setNotice(`The alert is saved. People see it within ${ALERT_CACHE_MINUTES} minutes.`);
        load();
      })
      .catch((err) => {
        if (mounted.current) {
          setSaving(false);
          setFormError(`The alert could not be saved: ${errorText(err)}`);
        }
      });
  }, [closeForm, end, headline, kit, link, load, message, mounted, saving, severity, start]);

  const removeAlert = useCallback(
    (alert: IAlert) => {
      kit.rootLists
        .recycleItem(ALERTS.urlName, alert.Id)
        .then(() => {
          if (!mounted.current) {
            return;
          }
          clearAlertCache(kit.sp);
          setAlerts((current) => (current ? current.filter((a) => a.Id !== alert.Id) : current));
          setNotice(`"${alert.Title}" was moved to the recycle bin. It disappears for everyone within ${ALERT_CACHE_MINUTES} minutes.`);
        })
        .catch((err) => {
          if (mounted.current) {
            setError(`The alert could not be removed: ${errorText(err)}`);
          }
        });
    },
    [kit, mounted]
  );

  if (!kit.canManage) {
    return (
      <Shell icon="Megaphone" title={props.properties.title || 'Alert Banner Settings'}>
        <Empty icon="Lock" text="Only site owners can manage the alert banner." />
      </Shell>
    );
  }

  return (
    <Shell
      icon="Megaphone"
      title={props.properties.title || 'Alert Banner Settings'}
      subtitle="Show an announcement at the top of every page in this site collection."
      error={error}
      onDismissError={() => setError(undefined)}
      notice={notice}
      onDismissNotice={() => setNotice(undefined)}
      actions={
        !needsSetup &&
        !loading && (
          <>
            <DefaultButton iconProps={{ iconName: 'Edit' }} text="Open the list" href={kit.rootLists.listUrl(ALERTS.urlName)} target="_blank" />
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="New alert" onClick={() => setFormOpen(true)} />
          </>
        )
      }
    >
      {loading && <Spinner size={SpinnerSize.medium} label="Loading..." />}

      {needsSetup && !loading && (
        <SetupCard canManage={kit.canManage} what={'a "Site Alerts" list on the top-level site'} onSetup={setup} onDone={load} />
      )}

      {!needsSetup && !loading && (
        <div className={styles.card}>
          <Toggle
            label="Banner"
            inlineLabel
            onText="On: active alerts show on every page"
            offText="Off: alerts are not shown anywhere"
            checked={!!actionId}
            disabled={switching || !bannerKnown}
            onChange={(_, checked) => switchBanner(!!checked)}
          />
          {!bannerKnown && (
            <div className={styles.muted}>Only a site collection administrator can see or change whether the banner is switched on.</div>
          )}
          <div className={styles.small}>
            The banner reads the alert list at most once every {ALERT_CACHE_MINUTES} minutes per person, so changes take up to that long to
            reach everyone.
          </div>
        </div>
      )}

      {!needsSetup && !loading && alerts && alerts.length === 0 && <Empty icon="Megaphone" text="There are no alerts yet." />}

      {!needsSetup && alerts && alerts.length > 0 && (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Alert</th>
                <th>Severity</th>
                <th>Status</th>
                <th>From</th>
                <th>Until</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {alerts.map((alert) => {
                const status = alertStatus(alert);
                return (
                  <tr key={alert.Id}>
                    <td>
                      <div className={styles.fileName}>{alert.Title}</div>
                      {alert.Message && <div className={styles.path}>{alert.Message}</div>}
                    </td>
                    <td>{alert.Severity || 'Info'}</td>
                    <td>
                      <span className={`${styles.pill} ${STATUS_STYLE[status]}`}>{status}</span>
                    </td>
                    <td className={styles.nowrap}>{formatDate(alert.StartDate) || 'Now'}</td>
                    <td className={styles.nowrap}>{formatDate(alert.EndDate) || 'No end'}</td>
                    <td>
                      <button type="button" className={styles.linkButton} onClick={() => removeAlert(alert)}>
                        Remove
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <Panel isOpen={formOpen} onDismiss={() => !saving && closeForm()} type={PanelType.smallFixedFar} headerText="New alert" closeButtonAriaLabel="Close">
        <div className={styles.stack}>
          <TextField label="Headline" required value={headline} maxLength={255} onChange={(_, v) => setHeadline(v || '')} disabled={saving} />
          <TextField label="Message" multiline rows={4} value={message} maxLength={500} onChange={(_, v) => setMessage(v || '')} disabled={saving} />
          <Dropdown
            label="Severity"
            options={SEVERITIES.map((s) => ({ key: s, text: s }))}
            selectedKey={severity}
            onChange={(_, option) => option && setSeverity(String(option.key))}
            disabled={saving}
          />
          <TextField label="Show from" type="date" value={start} onChange={(_, v) => setStart(v || '')} disabled={saving} />
          <TextField label="Show until (optional)" type="date" value={end} onChange={(_, v) => setEnd(v || '')} disabled={saving} />
          <TextField label="Link (optional)" placeholder="https://..." value={link} onChange={(_, v) => setLink(v || '')} disabled={saving} />
          {formError && <div style={{ color: '#a4262c' }}>{formError}</div>}
          <div className={styles.row}>
            <PrimaryButton text={saving ? 'Saving...' : 'Save alert'} disabled={saving} onClick={addAlert} />
            <DefaultButton text="Cancel" disabled={saving} onClick={closeForm} />
          </div>
        </div>
      </Panel>
    </Shell>
  );
};
