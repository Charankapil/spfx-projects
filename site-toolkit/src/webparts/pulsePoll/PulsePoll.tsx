import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { DefaultButton, Spinner, SpinnerSize } from '@fluentui/react';

import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { formatNumber } from '../../common/format';
import { field, IListDefinition, isMissingList } from '../../common/lists';
import { errorText } from '../../common/RequestQueue';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { SetupCard } from '../../common/ui/SetupCard';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { parseOptions, tally } from './poll';

export interface IPulsePollProps {
  title: string;
  question: string;
  options: string;
  pollKey: string;
}

const VOTES: IListDefinition = {
  urlName: 'ToolkitPollVotes',
  title: 'Poll Votes',
  description: 'One row per vote cast with the Pulse Poll web part.',
  fields: [field.text('PollKey', 'Poll Key', true)],
  indexAuthor: true,
  ownItemsOnly: true
};

const MAX_VOTES_COUNTED = 5000;

export const PulsePoll: React.FC<IToolkitComponentProps<IPulsePollProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const question = (props.properties.question || '').trim();
  const options = useMemo(() => parseOptions(props.properties.options), [props.properties.options]);
  // Each web part on a page is its own poll unless the owner gives it a key.
  const pollKey = ((props.properties.pollKey || '').trim() || props.context.instanceId).substring(0, 200);
  const filterKey = pollKey.replace(/'/g, "''");
  const myVoteKey = cacheKey(kit.scope, `poll:${pollKey}:vote:${kit.userId}`);
  const resultsKey = cacheKey(kit.scope, `poll:${pollKey}:results`);

  const [myVote, setMyVote] = useState<string | undefined>(undefined);
  const [counts, setCounts] = useState<Record<string, number> | undefined>(undefined);
  const [capped, setCapped] = useState(false);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [loading, setLoading] = useState(true);
  const [voting, setVoting] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const loadResults = useCallback(
    async (force: boolean): Promise<void> => {
      let votes = force ? undefined : cacheGet<{ votes: string[]; capped: boolean }>('session', resultsKey);
      if (!votes) {
        const all = await kit.lists.getItems<{ Title: string }>(
          VOTES.urlName,
          `$select=Title&$filter=PollKey eq '${encodeURIComponent(filterKey)}'&$top=2000`,
          MAX_VOTES_COUNTED
        );
        votes = { votes: all.items.map((v) => v.Title), capped: all.more };
        cacheSet('session', resultsKey, votes, 1);
      }
      if (mounted.current) {
        setCounts(tally(options, votes.votes));
        setCapped(votes.capped);
      }
    },
    [filterKey, kit, mounted, options, resultsKey]
  );

  const load = useCallback(
    (force: boolean) => {
      if (!question || options.length < 2) {
        setLoading(false);
        return;
      }
      setLoading(true);
      setError(undefined);
      const run = async (): Promise<string | undefined> => {
        // The browser remembers that this person voted, so most page loads ask nothing about it.
        let vote = cacheGet<string>('local', myVoteKey);
        if (!vote) {
          const mine = await kit.lists.getItems<{ Title: string }>(
            VOTES.urlName,
            `$select=Id,Title&$filter=PollKey eq '${encodeURIComponent(filterKey)}' and AuthorId eq ${kit.userId}&$top=1`,
            1
          );
          vote = mine.items.length > 0 ? mine.items[0].Title : undefined;
          if (vote) {
            cacheSet('local', myVoteKey, vote, 60 * 24 * 90);
          }
        }
        if (vote || kit.canManage) {
          await loadResults(force);
        }
        return vote;
      };
      run()
        .then((vote) => {
          if (mounted.current) {
            setMyVote(vote);
            setNeedsSetup(false);
            setLoading(false);
          }
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setLoading(false);
          if (isMissingList(err)) {
            setNeedsSetup(true);
          } else {
            setError(`The poll could not be loaded: ${errorText(err)}`);
          }
        });
    },
    [filterKey, kit, loadResults, mounted, myVoteKey, options.length, question]
  );

  useEffect(() => {
    load(false);
  }, [load]);

  const vote = useCallback(
    (choice: string) => {
      if (voting || myVote) {
        return;
      }
      setVoting(true);
      setError(undefined);
      kit.lists
        .addItem(VOTES.urlName, { Title: choice, PollKey: pollKey })
        .then(() => {
          cacheSet('local', myVoteKey, choice, 60 * 24 * 90);
          return loadResults(true);
        })
        .then(() => {
          if (mounted.current) {
            setMyVote(choice);
            setVoting(false);
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setVoting(false);
            setError(`Your vote could not be saved: ${errorText(err)}`);
          }
        });
    },
    [kit, loadResults, mounted, myVote, myVoteKey, pollKey, voting]
  );

  const total = counts ? options.reduce((sum, option) => sum + (counts[option] || 0), 0) : 0;
  const showResults = !!counts && (!!myVote || kit.canManage);

  return (
    <Shell
      icon="Poll"
      title={props.properties.title || 'Pulse Poll'}
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        showResults && <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh results" disabled={loading} onClick={() => load(true)} />
      }
    >
      {(!question || options.length < 2) && (
        <Empty icon="Poll" text="Edit this web part and enter a question with at least two answers to start the poll." />
      )}

      {needsSetup && !loading && (
        <SetupCard canManage={kit.canManage} what={'a "Poll Votes" list'} onSetup={() => kit.lists.ensure(VOTES)} onDone={() => load(true)} />
      )}

      {loading && <Spinner size={SpinnerSize.medium} label="Loading the poll..." />}

      {question && options.length >= 2 && !needsSetup && !loading && (
        <div className={styles.card}>
          <h3 className={styles.cardTitle} style={{ fontSize: 16 }}>
            {question}
          </h3>
          {!myVote && (
            <div className={styles.stack} role="group" aria-label={question}>
              {options.map((option) => (
                <button key={option} type="button" className={styles.option} disabled={voting} onClick={() => vote(option)}>
                  {option}
                </button>
              ))}
              {voting && <div className={styles.small}>Saving your vote...</div>}
            </div>
          )}
          {showResults && counts && (
            <div className={styles.stack} style={{ marginTop: myVote ? 0 : 12 }}>
              {!myVote && <div className={styles.small}>Results so far (only site owners see these before voting)</div>}
              {options.map((option) => {
                const count = counts[option] || 0;
                const percent = total > 0 ? Math.round((count / total) * 100) : 0;
                return (
                  <div key={option} className={styles.result}>
                    <div className={styles.resultFill} style={{ width: `${percent}%` }} />
                    <div className={styles.resultText}>
                      <span>
                        {option}
                        {myVote === option && <strong> · your answer</strong>}
                      </span>
                      <span className={styles.nowrap}>
                        {percent}% ({formatNumber(count)})
                      </span>
                    </div>
                  </div>
                );
              })}
              <div className={styles.small}>
                {formatNumber(total)} vote{total === 1 ? '' : 's'}
                {capped ? ` (only the first ${formatNumber(MAX_VOTES_COUNTED)} are counted)` : ''}. Results refresh when you press Refresh.
              </div>
            </div>
          )}
        </div>
      )}
    </Shell>
  );
};
