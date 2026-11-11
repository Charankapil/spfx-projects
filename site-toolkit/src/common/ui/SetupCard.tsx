import * as React from 'react';
import { useCallback, useState } from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';

import { errorText } from '../RequestQueue';
import { Empty } from './Shell';

export interface ISetupCardProps {
  canManage: boolean;
  /** What will be created, e.g. 'a "FAQ" list'. */
  what: string;
  onSetup: () => Promise<void>;
  onDone: () => void;
}

/**
 * Shown when the list a web part needs does not exist yet. Only site owners
 * get the button, because creating a list needs the Manage Lists permission.
 */
export const SetupCard: React.FC<ISetupCardProps> = (props) => {
  const { onSetup, onDone } = props;
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const run = useCallback(() => {
    setBusy(true);
    setError(undefined);
    onSetup()
      .then(() => {
        setBusy(false);
        onDone();
      })
      .catch((err) => {
        setBusy(false);
        setError(`Setup did not finish: ${errorText(err)}`);
      });
  }, [onSetup, onDone]);

  if (!props.canManage) {
    return <Empty icon="Settings" text="This web part has not been set up on this site yet. Please ask a site owner to open this page and set it up." />;
  }

  return (
    <Empty icon="Settings" text={`This web part stores its data in this site. Setting it up creates ${props.what}.`}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {busy ? <Spinner size={SpinnerSize.medium} label="Setting up..." /> : <PrimaryButton text="Set up" onClick={run} />}
    </Empty>
  );
};
