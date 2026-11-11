import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { DefaultButton, Icon, SearchBox, Spinner, SpinnerSize } from '@fluentui/react';

import { cacheGet, cacheKey, cacheRemove, cacheSet } from '../../common/cache';
import { field, IListDefinition, isMissingList } from '../../common/lists';
import { errorText } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { SetupCard } from '../../common/ui/SetupCard';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IFaqAccordionProps {
  title: string;
}

const FAQ: IListDefinition = {
  urlName: 'ToolkitFaq',
  title: 'FAQ',
  description: 'Questions and answers shown by the FAQ web part. The Title column is the question.',
  fields: [field.note('Answer', 'Answer'), field.text('Category', 'Category'), field.number('SortOrder', 'Sort Order')]
};

const SAMPLES = [
  {
    Title: 'How do I add a question here?',
    Answer: 'Site owners can press "Edit questions" on this web part, then add an item: the title is the question.',
    Category: 'Getting started',
    SortOrder: 10
  }
];

interface IFaq {
  Id: number;
  Title: string;
  Answer?: string;
  Category?: string;
  SortOrder?: number;
}

const ALL = '__all__';

export const FaqAccordion: React.FC<IToolkitComponentProps<IFaqAccordionProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const key = cacheKey(kit.scope, 'faq');

  const [items, setItems] = useState<IFaq[] | undefined>(undefined);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [text, setText] = useState('');
  const [category, setCategory] = useState<string>(ALL);
  const [open, setOpen] = useState<Record<number, boolean>>({});

  const load = useCallback(
    (force: boolean) => {
      setLoading(true);
      setError(undefined);
      const run = async (): Promise<IFaq[]> => {
        const hit = force ? undefined : cacheGet<IFaq[]>('session', key);
        if (hit) {
          return hit;
        }
        const result = await kit.lists.getItems<IFaq>(FAQ.urlName, '$select=Id,Title,Answer,Category,SortOrder&$top=500', 500);
        cacheSet('session', key, result.items, 15);
        return result.items;
      };
      run()
        .then((list) => {
          if (!mounted.current) {
            return;
          }
          setItems(list.slice().sort((a, b) => (a.SortOrder || 0) - (b.SortOrder || 0) || a.Id - b.Id));
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
            setError(`The questions could not be loaded: ${errorText(err)}`);
          }
        });
    },
    [key, kit, mounted]
  );

  useEffect(() => {
    load(false);
  }, [load]);

  const setup = useCallback(async () => {
    let existed = true;
    try {
      await kit.sp.getJson(`${kit.lists.api(FAQ.urlName)}?$select=Id`);
    } catch (err) {
      if (!isMissingList(err)) {
        throw err;
      }
      existed = false;
    }
    await kit.lists.ensure(FAQ);
    if (!existed) {
      for (const sample of SAMPLES) {
        await kit.lists.addItem(FAQ.urlName, sample);
      }
    }
    cacheRemove('session', key);
  }, [key, kit]);

  const categories = useMemo(() => {
    const seen: string[] = [];
    for (const item of items || []) {
      const name = (item.Category || '').trim();
      if (name && seen.indexOf(name) < 0) {
        seen.push(name);
      }
    }
    return seen.sort((a, b) => a.localeCompare(b));
  }, [items]);

  const visible = useMemo(() => {
    const needle = text.trim().toLowerCase();
    return (items || []).filter((item) => {
      if (category !== ALL && (item.Category || '').trim() !== category) {
        return false;
      }
      return !needle || `${item.Title} ${item.Answer || ''}`.toLowerCase().indexOf(needle) >= 0;
    });
  }, [category, items, text]);

  return (
    <Shell
      icon="Unknown"
      title={props.properties.title || 'FAQ'}
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        !needsSetup && (
          <>
            <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => load(true)} />
            {kit.canManage && (
              <DefaultButton iconProps={{ iconName: 'Edit' }} text="Edit questions" href={kit.lists.listUrl(FAQ.urlName)} target="_blank" />
            )}
          </>
        )
      }
    >
      {loading && <Spinner size={SpinnerSize.medium} label="Loading..." />}

      {needsSetup && !loading && (
        <SetupCard canManage={kit.canManage} what={'an "FAQ" list with one example question'} onSetup={setup} onDone={() => load(true)} />
      )}

      {!needsSetup && items && items.length > 0 && (
        <>
          <SearchBox placeholder="Search the questions and answers" value={text} onChange={(_, value) => setText(value || '')} />
          {categories.length > 1 && (
            <div className={styles.row} role="group" aria-label="Category">
              {[ALL].concat(categories).map((name) => (
                <DefaultButton
                  key={name}
                  text={name === ALL ? 'All' : name}
                  primary={category === name}
                  aria-pressed={category === name}
                  onClick={() => setCategory(name)}
                  styles={{ root: { minWidth: 0, height: 28, borderRadius: 14 } }}
                />
              ))}
            </div>
          )}
        </>
      )}

      {!needsSetup && !loading && items && visible.length === 0 && (
        <Empty icon="Unknown" text={items.length === 0 ? 'There are no questions yet.' : 'No question matches that search.'} />
      )}

      {!needsSetup && visible.length > 0 && (
        <div className={styles.stack}>
          {visible.map((item) => {
            const isOpen = !!open[item.Id];
            return (
              <div key={item.Id} className={styles.accordionItem}>
                <button
                  type="button"
                  className={styles.accordionHeader}
                  aria-expanded={isOpen}
                  aria-controls={`faq-answer-${item.Id}`}
                  onClick={() => setOpen((current) => ({ ...current, [item.Id]: !isOpen }))}
                >
                  <span>{item.Title}</span>
                  <Icon iconName={isOpen ? 'ChevronUp' : 'ChevronDown'} aria-hidden="true" />
                </button>
                {isOpen && (
                  <div id={`faq-answer-${item.Id}`} className={styles.accordionBody}>
                    {item.Answer || 'No answer has been written yet.'}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </Shell>
  );
};
