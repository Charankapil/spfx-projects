import * as React from 'react';
import { ChoiceGroup, DefaultButton, Dropdown, Icon, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { ICreateRequest, ICreateResult, ISiteTemplate, powerShellFor, SiteCreator, STATUS_TEXT, urlNameFromTitle, validAlias } from '../../services/SiteCreator';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, ErrorBar, Loading, Pill, ViewHeader } from '../shared/ui';

const POLL_MS = 10000;
const MAX_POLLS = 12;

/**
 * Create a site collection from any template: modern communication (Topic, Showcase,
 * Blank), team sites with or without a Microsoft 365 group, your organisation's
 * site templates, and classic templates (as a ready-to-run PowerShell command).
 */
export const CreateSiteView: React.FC = () => {
  const ctx = useAdmin();
  const creator = React.useMemo(() => new SiteCreator(ctx.client, ctx.homeWebUrl), [ctx.client, ctx.homeWebUrl]);
  const origin = (/^(https:\/\/[^/]+)/i.exec(ctx.homeWebUrl) || [''])[1];
  const loader = useLoader(async () => {
    const [templates, languages, hubs] = await Promise.all([creator.templates(), creator.languages(), creator.hubs()]);
    return { templates, languages, hubs };
  }, []);

  const [template, setTemplate] = React.useState<ISiteTemplate | undefined>();
  const [title, setTitle] = React.useState('');
  const [alias, setAlias] = React.useState('');
  const [aliasEdited, setAliasEdited] = React.useState(false);
  const [managedPath, setManagedPath] = React.useState('sites');
  const [description, setDescription] = React.useState('');
  const [lcid, setLcid] = React.useState(1033);
  const [owner, setOwner] = React.useState('');
  const [isPublic, setIsPublic] = React.useState(false);
  const [hubId, setHubId] = React.useState('');
  const [check, setCheck] = React.useState<{ ok: boolean; message: string; suggestion?: string } | undefined>();
  const [busy, setBusy] = React.useState<string | undefined>();
  const [result, setResult] = React.useState<(ICreateResult & { title: string; request: ICreateRequest; extra: string[] }) | undefined>();
  const [showPs, setShowPs] = React.useState(false);
  const alive = React.useRef(true);
  // Template and hub steps run once per created site, whether the poll or a manual check sees it ready first.
  const finished = React.useRef<{ [url: string]: boolean }>({});
  const finishOnce = async (req: ICreateRequest, res: ICreateResult): Promise<string[]> => {
    if (finished.current[res.url]) {
      return [];
    }
    finished.current[res.url] = true;
    return creator.finish(req, res);
  };
  React.useEffect(
    () => () => {
      alive.current = false;
    },
    []
  );

  if (loader.loading && !loader.data) {
    return <Loading text="Loading site templates…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const { templates, languages, hubs } = loader.data;
  const t = template || templates[0];
  const isGroup = t.route === 'group';
  const request: ICreateRequest = { template: t, title: title.trim(), managedPath: isGroup ? 'teams' : managedPath, alias: alias.trim(), description: description.trim(), lcid, owner: owner.trim(), isPublic, hubId: hubId || undefined };
  const url = `${origin}/${request.managedPath}/${request.alias || '…'}`;
  const formOk = !!request.title && validAlias(request.alias);

  const onTitle = (v: string): void => {
    setTitle(v);
    if (!aliasEdited) {
      setAlias(urlNameFromTitle(v));
    }
    setCheck(undefined);
  };

  const runCheck = async (): Promise<boolean> => {
    setBusy('Checking the address…');
    try {
      const c = await creator.checkAvailable(request);
      setCheck(c);
      return c.ok;
    } catch (e) {
      setCheck({ ok: false, message: (e as Error).message });
      return false;
    } finally {
      setBusy(undefined);
    }
  };

  const poll = async (res: ICreateResult, req: ICreateRequest, extra: string[]): Promise<void> => {
    for (let i = 0; i < MAX_POLLS && alive.current; i++) {
      await new Promise((r) => setTimeout(r, POLL_MS));
      if (!alive.current) {
        return;
      }
      try {
        res.status = await creator.refreshStatus(res);
      } catch {
        /* try again on the next round */
      }
      setResult({ ...res, title: req.title, request: req, extra });
      if (res.status === 2) {
        const more = await finishOnce(req, res);
        setResult({ ...res, title: req.title, request: req, extra: extra.concat(more) });
        ctx.notify(`${req.title} is ready.`, 'success');
        return;
      }
      if (res.status === 3) {
        return;
      }
    }
  };

  const create = async (): Promise<void> => {
    if (!(await runCheck())) {
      return;
    }
    const ok = await ctx.confirm({
      title: 'Create site',
      message: `Create "${request.title}" at ${url} using "${t.title}"${request.owner ? ` with ${request.owner} as owner` : ''}?`,
      confirmText: 'Create'
    });
    if (!ok) {
      return;
    }
    setBusy('Creating the site…');
    try {
      const res = await creator.create(request);
      ctx.log('Create site collection', `${res.url} (${t.title})`, true);
      const extra: string[] = [];
      setResult({ ...res, title: request.title, request, extra });
      if (res.status === 2) {
        const more = await finishOnce(request, res);
        setResult({ ...res, title: request.title, request, extra: more });
        ctx.notify(`${request.title} was created.`, 'success');
      } else {
        ctx.notify(`${request.title} is being created. This page checks its status every 10 seconds.`, 'info');
        poll(res, request, extra).catch(() => undefined);
      }
    } catch (e) {
      ctx.log('Create site collection', url, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    } finally {
      setBusy(undefined);
    }
  };

  const manage = async (siteUrl: string): Promise<void> => {
    try {
      const [web, site] = await Promise.all([ctx.api.getWeb(siteUrl), ctx.api.getSite(siteUrl)]);
      ctx.setTarget({ webUrl: web.url, siteUrl: site.url, title: web.title });
      ctx.openView('overview');
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  const ps = powerShellFor(request.alias ? request : { ...request, alias: 'url-name' }, origin, ctx.currentUser);
  const groups: Array<ISiteTemplate['group']> = ['Modern', 'Your organisation', 'Classic'];

  return (
    <div className={styles.view}>
      <ViewHeader title="Create site" hint="Create a site collection from any template, as yourself. Your tenant's site-creation settings and your rights apply." />

      <Card title="1. Choose a template">
        {groups.map((g) => {
          const list = templates.filter((x) => x.group === g);
          if (!list.length) {
            return null;
          }
          return (
            <React.Fragment key={g}>
              <div className={styles.templateGroup}>
                {g === 'Classic' ? 'Classic templates (created with PowerShell)' : g === 'Your organisation' ? 'Your organisation’s templates' : 'Modern templates'}
              </div>
              <div className={styles.templateGrid} role="radiogroup" aria-label={g + ' templates'}>
                {list.map((x) => (
                  <button
                    key={x.key}
                    type="button"
                    role="radio"
                    aria-checked={x.key === t.key}
                    className={`${styles.templateTile} ${x.key === t.key ? styles.templateTileActive : ''}`}
                    onClick={() => {
                      setTemplate(x);
                      setCheck(undefined);
                    }}
                  >
                    <Icon iconName={x.icon} aria-hidden="true" />
                    <span style={{ minWidth: 0 }}>
                      <div className={styles.templateTitle}>{x.title}</div>
                      <div className={styles.templateDesc}>{x.description}</div>
                    </span>
                  </button>
                ))}
              </div>
            </React.Fragment>
          );
        })}
      </Card>

      <div className={styles.grid}>
        <Card title="2. Site details">
          <TextField label="Site name" required value={title} onChange={(_, v) => onTitle(v || '')} maxLength={255} />
          <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end', flexWrap: 'wrap' }}>
            {!isGroup && (
              <Dropdown
                label="Path"
                selectedKey={managedPath}
                onChange={(_, o) => {
                  if (o) {
                    setManagedPath(String(o.key));
                    setCheck(undefined);
                  }
                }}
                options={[
                  { key: 'sites', text: '/sites/' },
                  { key: 'teams', text: '/teams/' }
                ]}
                styles={{ root: { width: 110 } }}
              />
            )}
            <TextField
              label={isGroup ? 'Group e-mail alias and URL name' : 'URL name'}
              required
              value={alias}
              onChange={(_, v) => {
                setAlias((v || '').replace(/\s+/g, ''));
                setAliasEdited(true);
                setCheck(undefined);
              }}
              errorMessage={alias && !validAlias(alias) ? 'Letters, numbers, - and _ only.' : undefined}
              styles={{ root: { flex: 1, minWidth: 180 } }}
            />
          </div>
          <p className={styles.muted} style={{ margin: '6px 0 0', fontSize: 12, wordBreak: 'break-all' }}>
            {url}
            {isGroup && ' (group sites use the tenant’s default path)'}
          </p>
          <TextField label="Description" value={description} onChange={(_, v) => setDescription(v || '')} multiline rows={2} maxLength={500} />
          {isGroup && (
            <ChoiceGroup
              label="Privacy"
              selectedKey={isPublic ? 'public' : 'private'}
              onChange={(_, o) => o && setIsPublic(o.key === 'public')}
              options={[
                { key: 'private', text: 'Private: only members can access' },
                { key: 'public', text: 'Public: anyone in the organisation can access' }
              ]}
            />
          )}
        </Card>
        <Card title="3. Options">
          <Dropdown label="Language" selectedKey={lcid} onChange={(_, o) => o && setLcid(Number(o.key))} options={languages.map((l) => ({ key: l.lcid, text: l.name }))} />
          <TextField label="Primary owner (optional)" placeholder={`Leave empty for you (${ctx.currentUser || 'signed-in user'})`} value={owner} onChange={(_, v) => setOwner(v || '')} />
          <Dropdown
            label="Join a hub (optional)"
            selectedKey={hubId}
            onChange={(_, o) => o && setHubId(String(o.key))}
            options={[{ key: '', text: 'No hub' }].concat(hubs.map((h) => ({ key: h.id, text: h.title })))}
          />
          <p className={styles.muted} style={{ fontSize: 12 }}>
            The language cannot be changed after creation. Setting another person as owner needs SharePoint admin rights.
          </p>
        </Card>
      </div>

      <Card title="4. Create">
        {t.route === 'powershell' && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline>
            Classic site collections can only be created by the tenant admin API (SharePoint admin center or PowerShell), which a web part cannot call. Run the command below as a SharePoint Administrator.
          </MessageBar>
        )}
        {check && (
          <MessageBar messageBarType={check.ok ? MessageBarType.success : MessageBarType.warning} isMultiline>
            {check.message}
            {check.suggestion && (
              <>
                {' '}
                <button
                  type="button"
                  className={styles.link}
                  onClick={() => {
                    setAlias(check.suggestion as string);
                    setAliasEdited(true);
                    setCheck(undefined);
                  }}
                >
                  Use &ldquo;{check.suggestion}&rdquo;
                </button>
              </>
            )}
          </MessageBar>
        )}
        <div className={styles.actions} style={{ marginTop: 8 }}>
          <DefaultButton iconProps={{ iconName: 'CheckMark' }} onClick={runCheck} disabled={!formOk || !!busy}>
            Check address
          </DefaultButton>
          <PrimaryButton iconProps={{ iconName: 'Add' }} onClick={create} disabled={!formOk || !!busy || t.route === 'powershell'}>
            {busy || 'Create site'}
          </PrimaryButton>
          <DefaultButton iconProps={{ iconName: 'CommandPrompt' }} onClick={() => setShowPs(!showPs)}>
            {showPs || t.route === 'powershell' ? 'PowerShell command' : 'Show PowerShell'}
          </DefaultButton>
        </div>
        {(showPs || t.route === 'powershell') && (
          <>
            <pre className={styles.code}>{ps}</pre>
            <DefaultButton
              iconProps={{ iconName: 'Copy' }}
              style={{ marginTop: 6 }}
              onClick={() => {
                try {
                  navigator.clipboard.writeText(ps).then(
                    () => ctx.notify('Copied.', 'success'),
                    () => ctx.notify('Copy failed; select the text instead.', 'error')
                  );
                } catch {
                  ctx.notify('Copy failed; select the text instead.', 'error');
                }
              }}
            >
              Copy
            </DefaultButton>
          </>
        )}
      </Card>

      {result && (
        <Card title="Result">
          <p style={{ margin: 0 }}>
            <strong>{result.title}</strong>{' '}
            <Pill kind={result.status === 2 ? 'good' : result.status === 3 ? 'critical' : 'info'}>{STATUS_TEXT[result.status] || 'Unknown'}</Pill>
          </p>
          <p className={styles.muted} style={{ wordBreak: 'break-all' }}>
            {result.url}
          </p>
          {result.extra.map((x) => (
            <p key={x} className={styles.muted} style={{ margin: '4px 0' }}>
              {x}
            </p>
          ))}
          <div className={styles.actions}>
            <DefaultButton iconProps={{ iconName: 'OpenInNewWindow' }} href={result.url} target="_blank" disabled={result.status !== 2}>
              Open site
            </DefaultButton>
            <PrimaryButton iconProps={{ iconName: 'Settings' }} onClick={() => manage(result.url)} disabled={result.status !== 2}>
              Manage in this tool
            </PrimaryButton>
            {result.status === 1 && (
              <DefaultButton
                iconProps={{ iconName: 'Refresh' }}
                onClick={async () => {
                  result.status = await creator.refreshStatus(result).catch(() => result.status);
                  const more = result.status === 2 ? await finishOnce(result.request, result) : [];
                  setResult({ ...result, extra: result.extra.concat(more) });
                }}
              >
                Check status
              </DefaultButton>
            )}
          </div>
        </Card>
      )}
    </div>
  );
};
