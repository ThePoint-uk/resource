import * as React from 'react';
import { describeError } from '../services/GraphService';

export interface IAsyncState<T> {
  data: T | undefined;
  loading: boolean;
  error: string | undefined;
  reload: () => void;
  setData: (updater: (prev: T | undefined) => T | undefined) => void;
}

/**
 * Runs an async factory whenever `deps` change, exposing loading/error state.
 * Guards against state updates after unmount or after a newer request started.
 */
export function useAsync<T>(factory: () => Promise<T>, deps: React.DependencyList): IAsyncState<T> {
  const [data, setDataState] = React.useState<T | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [tick, setTick] = React.useState<number>(0);
  const requestRef = React.useRef<number>(0);

  React.useEffect(() => {
    const requestId: number = ++requestRef.current;
    setLoading(true);
    setError(undefined);
    factory()
      .then(result => {
        if (requestId === requestRef.current) {
          setDataState(result);
          setLoading(false);
        }
      })
      .catch((e: unknown) => {
        if (requestId === requestRef.current) {
          setError(describeError(e));
          setLoading(false);
        }
      });
    return () => {
      requestRef.current++;
    };
  }, [...deps, tick]);

  const reload = React.useCallback((): void => setTick(t => t + 1), []);
  const setData = React.useCallback((updater: (prev: T | undefined) => T | undefined): void => setDataState(prev => updater(prev)), []);

  return { data, loading, error, reload, setData };
}
