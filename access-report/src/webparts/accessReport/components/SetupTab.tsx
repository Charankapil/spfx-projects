import * as React from 'react';
import { DefaultButton, Dropdown, IconButton, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { IFound } from '../services/PeoplePicker';
import { EVERYONE_LOGIN, everyoneExceptExternalLogin, ISubject, ManualType, manualSubject, normaliseLogin, subjectKind } from '../services/permissions';
import { folderProblem, ISettings } from '../services/Store';
import styles from './AccessReport.module.scss';
import { IConfig, IServices } from './AccessReport';
import { Card, Empty, errorText } from './ui';

/** Who the report is about and where it looks - the configuration the monthly flow reads too. */
export const SetupTab: React.FC<{ services: IServices; config: IConfig; onChange: (c: IConfig) => void }> = ({ services, config, onChange }) => {
  const [busy, setBusy] = React.useState(false);
  const [message, setMessage] = React.useState<{ kind: MessageBarType; text: string } | undefined>();
  const [search, setSearch] = React.useState('');
  const [found, setFound] = React.useState<IFound[] | undefined>();
  const [manualType, setManualType] = React.useState<ManualType>('person');
  const [manualValue, setManualValue] = React.useState('');
  const [manualLabel, setManualLabel] = React.useState('');
  const [removing, setRemoving] = React.useState<number | undefined>();
  const [draft, setDraft] = React.useState<ISettings>(config.settings);

  const guard = async (work: () => Promise<void>): Promise<void> => {
    setBusy(true);
    setMessage(undefined);
    try {
      await work();
    } catch (e) {
      setMessage({ kind: MessageBarType.error, text: errorText(e) });
    } finally {
      setBusy(false);
    }
  };
  const reloadSubjects = async (): Promise<void> => onChange({ settings: config.settings, subjects: await services.store.subjects() });
  const has = (login: string): boolean => config.subjects.some((s) => s.login.toLowerCase() === normaliseLogin(login).toLowerCase());

  const add = (s: ISubject): Promise<void> =>
    guard(async () => {
      if (has(s.login)) {
        setMessage({ kind: MessageBarType.warning, text: `${s.label} is already in the list.` });
        return;
      }
      await services.store.addSubject(s);
      await reloadSubjects();
    });

  const find = (): Promise<void> =>
    guard(async () => {
      setFound(await services.picker.search(search));
    });

  const addManual = (): void => {
    const s = manualSubject(manualType, manualValue, manualLabel);
    if (typeof s === 'string') {
      setMessage({ kind: MessageBarType.error, text: s });
      return;
    }
    add(s)
      .then(() => {
        setManualValue('');
        setManualLabel('');
      })
      .catch(() => undefined);
  };

  const problem = folderProblem(draft.reportFolder, services.store.webPath);
  const dirty = draft.includeLibraries !== config.settings.includeLibraries || draft.siteFilter.trim() !== config.settings.siteFilter || draft.reportFolder.trim() !== config.settings.reportFolder;
  const saveSettings = (): Promise<void> =>
    guard(async () => {
      const folder = draft.reportFolder.trim().replace(/\/+$/, '') || services.store.defaultFolder;
      if (!(await services.store.folderExists(folder))) {
        setMessage({ kind: MessageBarType.error, text: `${folder} does not exist. Create the folder in a document library first, then save.` });
        return;
      }
      await services.store.saveSettings({ ...draft, reportFolder: folder });
      const settings = await services.store.settings();
      setDraft(settings);
      onChange({ settings, subjects: config.subjects });
      setMessage({ kind: MessageBarType.success, text: 'Saved. The next run - from this page or from the monthly flow - uses these options.' });
    });

  return (
    <div className={styles.tab}>
      {message && (
        <MessageBar messageBarType={message.kind} isMultiline onDismiss={() => setMessage(undefined)}>
          {message.text}
        </MessageBar>
      )}

      <Card title={`People and groups in the report (${config.subjects.filter((s) => s.active).length} of ${config.subjects.length} included)`}>
        {config.subjects.length === 0 ? (
          <Empty text="Nobody yet. Find a person or a group below and add them." />
        ) : (
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Name in the report</th>
                  <th>Kind</th>
                  <th>SharePoint login</th>
                  <th>Include</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {config.subjects.map((s) => (
                  <tr key={s.id}>
                    <td>{s.label}</td>
                    <td>
                      <span className={styles.pill}>{subjectKind(s.login)}</span>
                    </td>
                    <td className={styles.code}>{s.login}</td>
                    <td>
                      <Toggle
                        ariaLabel={`Include ${s.label}`}
                        checked={s.active}
                        disabled={busy}
                        styles={{ root: { margin: 0 } }}
                        onChange={(_, on) => guard(async () => services.store.setSubjectActive(s.id, !!on).then(reloadSubjects))}
                      />
                    </td>
                    <td style={{ whiteSpace: 'nowrap', textAlign: 'right' }}>
                      {removing === s.id ? (
                        <>
                          <DefaultButton
                            text="Remove"
                            disabled={busy}
                            onClick={() =>
                              guard(async () => {
                                await services.store.removeSubject(s.id);
                                setRemoving(undefined);
                                await reloadSubjects();
                              })
                            }
                          />{' '}
                          <DefaultButton text="Keep" onClick={() => setRemoving(undefined)} />
                        </>
                      ) : (
                        <IconButton iconProps={{ iconName: 'Delete' }} title={`Remove ${s.label}`} ariaLabel={`Remove ${s.label}`} disabled={busy} onClick={() => setRemoving(s.id)} />
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      <Card title="Add a person or a group">
        <p className={styles.hint}>Search the directory the way the Share dialog does: people (guests too), security groups and Microsoft 365 groups.</p>
        <div className={styles.actions}>
          <TextField
            label="Name or e-mail"
            value={search}
            onChange={(_, v) => setSearch(v || '')}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && search.trim() && !busy) {
                find().catch(() => undefined);
              }
            }}
            styles={{ root: { width: 320, maxWidth: '100%' } }}
          />
          <PrimaryButton onClick={find} disabled={busy || !search.trim()}>
            Find
          </PrimaryButton>
          {services.tenantId && (
            <DefaultButton
              onClick={() => add({ login: everyoneExceptExternalLogin(services.tenantId), label: 'Everyone except external users' })}
              disabled={busy || has(everyoneExceptExternalLogin(services.tenantId))}
            >
              Add “Everyone except external users”
            </DefaultButton>
          )}
          <DefaultButton onClick={() => add({ login: EVERYONE_LOGIN, label: 'Everyone' })} disabled={busy || has(EVERYONE_LOGIN)}>
            Add “Everyone”
          </DefaultButton>
        </div>
        {found &&
          (found.length === 0 ? (
            <Empty text="Nothing found. Try the full e-mail address, or add it by hand below." />
          ) : (
            <ul className={styles.found}>
              {found.map((f) => (
                <li key={f.login}>
                  <span style={{ minWidth: 0 }}>
                    {f.label} <span className={styles.pill}>{f.kind}</span>
                    <div className={`${styles.cellSub} ${styles.code}`}>{f.login}</div>
                  </span>
                  <DefaultButton onClick={() => add(f)} disabled={busy || has(f.login)}>
                    {has(f.login) ? 'Added' : 'Add'}
                  </DefaultButton>
                </li>
              ))}
            </ul>
          ))}

        <h3 style={{ marginTop: 18 }}>Or add by hand</h3>
        <div className={styles.actions}>
          <Dropdown
            label="Type"
            selectedKey={manualType}
            options={[
              { key: 'person', text: 'Person (e-mail)' },
              { key: 'security', text: 'Security group (object id)' },
              { key: 'm365', text: 'Microsoft 365 group (object id)' },
              { key: 'claim', text: 'SharePoint claim' }
            ]}
            onChange={(_, o) => o && setManualType(o.key as ManualType)}
            styles={{ root: { width: 230 } }}
          />
          <TextField
            label={manualType === 'person' ? 'E-mail address' : manualType === 'claim' ? 'Claim' : 'Object id'}
            placeholder={manualType === 'person' ? 'name@contoso.com' : manualType === 'claim' ? 'c:0t.c|tenant|…' : '00000000-0000-0000-0000-000000000000'}
            value={manualValue}
            onChange={(_, v) => setManualValue(v || '')}
            styles={{ root: { width: 320, maxWidth: '100%' } }}
          />
          <TextField label="Name in the report (optional)" value={manualLabel} onChange={(_, v) => setManualLabel(v || '')} styles={{ root: { width: 240, maxWidth: '100%' } }} />
          <DefaultButton onClick={addManual} disabled={busy || !manualValue.trim()}>
            Add
          </DefaultButton>
        </div>
      </Card>

      <Card title="Where to look, and where the report goes">
        <Toggle
          label="Include libraries"
          inlineLabel
          checked={draft.includeLibraries}
          onChange={(_, on) => setDraft({ ...draft, includeLibraries: !!on })}
          onText="Sites and their document libraries"
          offText="Sites only (faster)"
        />
        <div className={styles.actions}>
          <TextField
            label="Only sites whose title or address contains (optional)"
            value={draft.siteFilter}
            onChange={(_, v) => setDraft({ ...draft, siteFilter: v || '' })}
            styles={{ root: { width: 340, maxWidth: '100%' } }}
          />
          <TextField
            label="Report folder (server-relative path)"
            value={draft.reportFolder}
            onChange={(_, v) => setDraft({ ...draft, reportFolder: v || '' })}
            errorMessage={problem}
            styles={{ root: { width: 400, maxWidth: '100%' } }}
          />
        </div>
        <p className={`${styles.muted} ${styles.small}`}>
          The folder must be in this site ({services.store.webPath || '/'}), because that is where the monthly flow saves. Leave it empty for the default, {services.store.defaultFolder}. Whoever can
          open the folder can read the reports: give it the permissions the reports deserve.
        </p>
        <div className={styles.actions}>
          <PrimaryButton onClick={saveSettings} disabled={busy || !dirty || !!problem}>
            Save
          </PrimaryButton>
          {dirty && (
            <DefaultButton onClick={() => setDraft(config.settings)} disabled={busy}>
              Undo changes
            </DefaultButton>
          )}
        </div>
      </Card>
    </div>
  );
};
