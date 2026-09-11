import * as React from 'react';
import type { GraphService } from '../services/GraphService';

/** Resolves a user's photo through Graph, returning an object URL or undefined. */
export function usePhoto(graph: GraphService, userId: string | undefined, size: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' = '96x96'): string | undefined {
  const [url, setUrl] = React.useState<string | undefined>(undefined);
  React.useEffect(() => {
    let active: boolean = true;
    if (!userId) {
      setUrl(undefined);
      return;
    }
    graph.getPhoto(userId, size).then(result => {
      if (active) {
        setUrl(result);
      }
    }).catch(() => undefined);
    return () => {
      active = false;
    };
  }, [graph, userId, size]);
  return url;
}
