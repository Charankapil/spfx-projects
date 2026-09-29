import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { format } from './text';

interface IErrorBoundaryState {
  error?: Error;
}

/**
 * Keeps a rendering problem in one part of the dashboard from blanking the
 * whole web part (and the rest of the page's layout with it).
 */
export class ErrorBoundary extends React.Component<{ resetKey?: unknown }, IErrorBoundaryState> {
  public state: IErrorBoundaryState = {};

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error): void {
    console.error('[Storage Pulse] Rendering failed', error);
  }

  public componentDidUpdate(previous: { resetKey?: unknown }): void {
    if (this.state.error && previous.resetKey !== this.props.resetKey) {
      this.setState({ error: undefined });
    }
  }

  public render(): React.ReactNode {
    if (this.state.error) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          <strong>{strings.ErrorBoundaryTitle}</strong>{' '}
          {format(strings.ErrorBoundaryText, { error: this.state.error.message || String(this.state.error) })}
        </MessageBar>
      );
    }
    return this.props.children;
  }
}
