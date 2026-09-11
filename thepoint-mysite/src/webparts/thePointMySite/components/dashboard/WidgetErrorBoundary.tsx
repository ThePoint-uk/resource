import * as React from 'react';
import { ErrorState } from '../shared/States';

interface IState {
  error: Error | undefined;
}

/** Keeps one misbehaving widget from taking down the whole dashboard. */
export class WidgetErrorBoundary extends React.Component<{ resetKey?: string; children?: React.ReactNode }, IState> {
  public state: IState = { error: undefined };

  public static getDerivedStateFromError(error: Error): IState {
    return { error };
  }

  public componentDidUpdate(prev: { resetKey?: string }): void {
    if (prev.resetKey !== this.props.resetKey && this.state.error) {
      this.setState({ error: undefined });
    }
  }

  public componentDidCatch(error: Error): void {
    console.error('[ThePoint MySite] Widget crashed', error);
  }

  public render(): React.ReactNode {
    if (this.state.error) {
      return <ErrorState message={this.state.error.message} onRetry={() => this.setState({ error: undefined })} />;
    }
    return this.props.children;
  }
}
