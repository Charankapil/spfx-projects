import * as React from 'react';
import { Breadcrumb, Icon, MessageBar, MessageBarType, Pivot, PivotItem, Spinner } from '@fluentui/react';

import { extension, fileName, IFolderEntry, ILibraryInfo, WORD_EXTENSIONS } from '../services/DocumentService';
import { ESignService } from '../services/ESignService';
import { shortId } from '../services/hash';
import { errorText, formatDate } from './shared';
import styles from './ESign.module.scss';

export interface IPickedDocument {
  name: string;
  /** Where it came from in SharePoint (absent for uploads). */
  sourceUrl?: string;
  /** Always PDF - Word files are converted on the way in. */
  pdf: ArrayBuffer;
  converted: boolean;
}

export interface IDocumentPickerProps {
  service: ESignService;
  initialUrl?: string;
  onPicked: (doc: IPickedDocument) => void;
}

const MAX_BYTES = 50 * 1024 * 1024;

export function DocumentPicker(props: IDocumentPickerProps): JSX.Element {
  const { service } = props;
  const [libraries, setLibraries] = React.useState<ILibraryInfo[]>();
  const [path, setPath] = React.useState<{ name: string; url: string }[]>([]);
  const [entries, setEntries] = React.useState<IFolderEntry[]>();
  const [busy, setBusy] = React.useState<string>();
  const [error, setError] = React.useState<string>();

  const load = async (name: string, url: string, sourceUrl?: string): Promise<void> => {
    setError(undefined);
    const ext = extension(name);
    try {
      let pdf: ArrayBuffer;
      let converted = false;
      if (WORD_EXTENSIONS.indexOf(ext) >= 0) {
        setBusy(`Converting ${name} to PDF...`);
        pdf = await service.docs.convertToPdf(url);
        converted = true;
      } else if (ext === 'pdf') {
        setBusy(`Opening ${name}...`);
        pdf = await service.docs.read(url);
      } else {
        throw new Error('Choose a PDF or Word document.');
      }
      props.onPicked({ name, sourceUrl: sourceUrl || url, pdf, converted });
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };

  React.useEffect(() => {
    service.docs.libraries().then(setLibraries).catch(e => setError(errorText(e)));
    if (props.initialUrl) {
      load(fileName(props.initialUrl), props.initialUrl).catch(() => undefined);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const open = async (crumbs: { name: string; url: string }[]): Promise<void> => {
    setPath(crumbs);
    setEntries(undefined);
    if (crumbs.length) {
      try {
        setEntries(await service.docs.folder(crumbs[crumbs.length - 1].url));
      } catch (e) {
        setError(errorText(e));
      }
    }
  };

  const upload = async (file: File | undefined): Promise<void> => {
    if (!file) {
      return;
    }
    setError(undefined);
    if (file.size > MAX_BYTES) {
      setError('That file is larger than 50 MB.');
      return;
    }
    const ext = extension(file.name);
    const content = await file.arrayBuffer();
    if (ext === 'pdf') {
      props.onPicked({ name: file.name, pdf: content, converted: false });
      return;
    }
    if (WORD_EXTENSIONS.indexOf(ext) < 0) {
      setError('Choose a PDF or Word document.');
      return;
    }
    try {
      setBusy(`Converting ${file.name} to PDF...`);
      // SharePoint can only convert a file it holds, so stage the upload in the eSignature library first.
      const staging = `${service.libraryUrl}/_staging-${shortId()}`;
      await service.docs.ensureFolder(staging);
      const pdf = await service.docs.convertUploadToPdf(staging, file.name, content);
      props.onPicked({ name: file.name, pdf, converted: true });
    } catch (e) {
      setError(`${errorText(e)} - save the document as PDF and upload that instead.`);
    } finally {
      setBusy(undefined);
    }
  };

  return (
    <div>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      {busy && <Spinner label={busy} styles={{ root: { margin: 16 } }} />}
      <Pivot>
        <PivotItem headerText="From this site" itemIcon="SharepointLogo">
          <Breadcrumb
            items={[{ text: 'Libraries', key: 'root', onClick: () => open([]) }].concat(
              path.map((p, i) => ({ text: p.name, key: p.url, onClick: () => open(path.slice(0, i + 1)) }))
            )}
          />
          <div className={styles.browser}>
            {!path.length &&
              (libraries ? (
                libraries.map(l => (
                  <button key={l.url} className={styles.browserRow} onClick={() => open([{ name: l.title, url: l.url }])}>
                    <Icon iconName="DocLibrary" /> {l.title}
                  </button>
                ))
              ) : (
                <Spinner styles={{ root: { padding: 16 } }} />
              ))}
            {path.length > 0 && !entries && <Spinner styles={{ root: { padding: 16 } }} />}
            {path.length > 0 && entries && entries.length === 0 && <div className={styles.empty}>No PDF or Word documents here.</div>}
            {path.length > 0 &&
              entries &&
              entries.map(e => (
                <button
                  key={e.url}
                  className={styles.browserRow}
                  disabled={!!busy}
                  onClick={() => (e.isFolder ? open(path.concat({ name: e.name, url: e.url })) : load(e.name, e.url))}
                >
                  <Icon iconName={e.isFolder ? 'FabricFolder' : extension(e.name) === 'pdf' ? 'PDF' : 'WordDocument'} />
                  <span style={{ flex: 1 }}>{e.name}</span>
                  {!e.isFolder && <span className={styles.muted}>{formatDate(e.modified)}</span>}
                </button>
              ))}
          </div>
        </PivotItem>
        <PivotItem headerText="Upload from computer" itemIcon="Upload">
          <label className={styles.dropzone} style={{ display: 'block', marginTop: 12 }}>
            <Icon iconName="CloudUpload" style={{ fontSize: 28 }} />
            <div>Choose a PDF or Word document (up to 50 MB)</div>
            <input
              type="file"
              accept=".pdf,.doc,.docx,.docm,.dotx,.odt,.rtf"
              style={{ display: 'none' }}
              onChange={e => upload(e.target.files?.[0])}
            />
          </label>
        </PivotItem>
      </Pivot>
    </div>
  );
}
