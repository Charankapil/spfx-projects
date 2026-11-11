import * as React from 'react';
import { useCallback, useEffect, useState } from 'react';
import { DefaultButton, Dropdown, Icon, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, Spinner, SpinnerSize, TextField } from '@fluentui/react';

import { cacheGet, cacheKey, cacheRemove, cacheSet } from '../../common/cache';
import { extensionOf, fileIcon, formatDate } from '../../common/format';
import { getLibraries, ILibrary } from '../../common/libraries';
import { isMissingList } from '../../common/lists';
import { errorText } from '../../common/RequestQueue';
import { quote } from '../../common/SpClient';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { SetupCard } from '../../common/ui/SetupCard';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { stem, validateFileName } from './fileNames';

export interface ITemplateGalleryProps {
  title: string;
  libraryTitle: string;
}

interface ITemplate {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Modified: string;
}

export const TemplateGallery: React.FC<IToolkitComponentProps<ITemplateGalleryProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const libraryTitle = (props.properties.libraryTitle || 'Templates').trim();
  const key = cacheKey(kit.scope, `templates:${libraryTitle}`);

  const [templates, setTemplates] = useState<ITemplate[] | undefined>(undefined);
  const [libraries, setLibraries] = useState<ILibrary[]>([]);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [created, setCreated] = useState<{ name: string; url: string } | undefined>(undefined);

  const [chosen, setChosen] = useState<ITemplate | undefined>(undefined);
  const [name, setName] = useState('');
  const [target, setTarget] = useState<string | undefined>(undefined);
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState<string | undefined>(undefined);

  const load = useCallback(
    (force: boolean) => {
      setLoading(true);
      setError(undefined);
      const run = async (): Promise<{ templates: ITemplate[]; libraries: ILibrary[] }> => {
        let list = force ? undefined : cacheGet<ITemplate[]>('session', key);
        if (!list) {
          const result = await kit.sp.getAll<ITemplate>(
            `${kit.lists.libraryApi(libraryTitle)}/items?$select=Id,FileLeafRef,FileRef,Modified&$filter=FSObjType eq 0&$top=200`,
            200
          );
          list = result.items;
          cacheSet('session', key, list, 15);
        }
        const all = await getLibraries(kit.sp, kit.scope);
        return { templates: list, libraries: all };
      };
      run()
        .then((result) => {
          if (!mounted.current) {
            return;
          }
          result.templates.sort((a, b) => a.FileLeafRef.localeCompare(b.FileLeafRef));
          setTemplates(result.templates);
          // New documents never go into the template library itself.
          const targets = result.libraries.filter((l) => l.title.toLowerCase() !== libraryTitle.toLowerCase());
          setLibraries(targets);
          setTarget((current) => current || (targets.length > 0 ? targets[0].id : undefined));
          setNeedsSetup(false);
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setLoading(false);
          if (isMissingList(err)) {
            setNeedsSetup(true);
          } else {
            setError(`The templates could not be loaded: ${errorText(err)}`);
          }
        });
    },
    [key, kit, libraryTitle, mounted]
  );

  useEffect(() => {
    load(false);
  }, [load]);

  const setup = useCallback(async () => {
    await kit.lists.ensureLibrary(libraryTitle, 'Approved templates offered by the Template Gallery web part.');
    cacheRemove('session', key);
    cacheRemove('session', cacheKey(kit.scope, 'libraries'));
  }, [key, kit, libraryTitle]);

  const choose = useCallback((template: ITemplate) => {
    setChosen(template);
    setName(stem(template.FileLeafRef));
    setFormError(undefined);
  }, []);

  const create = useCallback(() => {
    const library = libraries.filter((l) => l.id === target)[0];
    const problem = validateFileName(name);
    if (!chosen || !library || saving) {
      return;
    }
    if (problem) {
      setFormError(problem);
      return;
    }
    const extension = extensionOf(chosen.FileLeafRef);
    const fileName = extension ? `${name.trim()}.${extension}` : name.trim();
    const destination = `${library.path}/${fileName}`;
    setSaving(true);
    setFormError(undefined);
    kit.sp
      .post(`${kit.sp.fileApi(chosen.FileRef)}/copyTo(strNewUrl='${quote(destination)}',bOverwrite=false)`)
      .then(() => {
        if (!mounted.current) {
          return;
        }
        setSaving(false);
        setChosen(undefined);
        setCreated({ name: fileName, url: `${encodeURI(destination).replace(/#/g, '%23').replace(/\?/g, '%3F')}?web=1` });
      })
      .catch((err) => {
        if (mounted.current) {
          setSaving(false);
          setFormError(`The document could not be created: ${errorText(err)}`);
        }
      });
  }, [chosen, kit, libraries, mounted, name, saving, target]);

  return (
    <Shell
      icon="DocumentSet"
      title={props.properties.title || 'Template Gallery'}
      subtitle="Start a new document from an approved template."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={!needsSetup && <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => load(true)} />}
    >
      {created && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setCreated(undefined)} dismissButtonAriaLabel="Close">
          &quot;{created.name}&quot; is ready.{' '}
          <a className={styles.link} href={created.url} target="_blank" rel="noopener noreferrer">
            Open it
          </a>
        </MessageBar>
      )}

      {loading && <Spinner size={SpinnerSize.medium} label="Loading templates..." />}

      {needsSetup && !loading && (
        <SetupCard canManage={kit.canManage} what={`a "${libraryTitle}" document library to hold the templates`} onSetup={setup} onDone={() => load(true)} />
      )}

      {!needsSetup && !loading && templates && templates.length === 0 && (
        <Empty icon="DocumentSet" text={`There are no templates yet. Upload documents to the "${libraryTitle}" library to offer them here.`} />
      )}

      {!needsSetup && templates && templates.length > 0 && (
        <div className={styles.grid}>
          {templates.map((template) => (
            <div key={template.Id} className={styles.tile}>
              <Icon iconName={fileIcon(extensionOf(template.FileLeafRef))} className={styles.tileIcon} aria-hidden="true" />
              <div className={styles.tileTitle}>{stem(template.FileLeafRef)}</div>
              <div className={styles.small}>
                {extensionOf(template.FileLeafRef).toUpperCase()} · updated {formatDate(template.Modified)}
              </div>
              <DefaultButton text="Use this template" onClick={() => choose(template)} disabled={libraries.length === 0} />
            </div>
          ))}
        </div>
      )}
      {!needsSetup && templates && templates.length > 0 && libraries.length === 0 && !loading && (
        <div className={styles.muted}>This site has no other document library to create the new document in.</div>
      )}

      <Panel
        isOpen={!!chosen}
        onDismiss={() => !saving && setChosen(undefined)}
        type={PanelType.smallFixedFar}
        headerText="New document"
        closeButtonAriaLabel="Close"
      >
        {chosen && (
          <div className={styles.stack}>
            <div className={styles.muted}>From the template &quot;{stem(chosen.FileLeafRef)}&quot;</div>
            <TextField
              label="Name"
              value={name}
              onChange={(_, value) => setName(value || '')}
              suffix={extensionOf(chosen.FileLeafRef) ? `.${extensionOf(chosen.FileLeafRef)}` : undefined}
              disabled={saving}
            />
            <Dropdown
              label="Save in"
              options={libraries.map((l) => ({ key: l.id, text: l.title }))}
              selectedKey={target}
              onChange={(_, option) => option && setTarget(String(option.key))}
              disabled={saving}
            />
            {formError && <div style={{ color: '#a4262c' }}>{formError}</div>}
            <div className={styles.row}>
              <PrimaryButton text={saving ? 'Creating...' : 'Create'} disabled={saving || !target} onClick={create} />
              <DefaultButton text="Cancel" disabled={saving} onClick={() => setChosen(undefined)} />
            </div>
          </div>
        )}
      </Panel>
    </Shell>
  );
};
