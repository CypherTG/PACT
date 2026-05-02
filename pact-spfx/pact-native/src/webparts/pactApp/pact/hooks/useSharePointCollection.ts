import React from 'react';

type Loader<T> = () => Promise<T[]>;

interface UseSharePointCollectionOptions {
  intervalMs?: number;
  refreshOnFocus?: boolean;
  refreshOnVisible?: boolean;
}

export function useSharePointCollection<T>(
  loader: Loader<T>,
  options: UseSharePointCollectionOptions = {}
): {
  data: T[];
  loading: boolean;
  error: Error | undefined;
  refresh: () => Promise<void>;
} {
  const { intervalMs = 0, refreshOnFocus = false, refreshOnVisible = false } = options;
  const [data, setData] = React.useState<T[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  // Store loader in a ref so it never triggers re-renders or effect re-runs
  const loaderRef = React.useRef(loader);
  loaderRef.current = loader;

  const refresh = React.useCallback(async (): Promise<void> => {
    try {
      const result = await loaderRef.current();
      setData(Array.isArray(result) ? result : []);
      setError(undefined);
    } catch (err) {
      setError(err instanceof Error ? err : new Error('SharePoint collection load failed'));
    } finally {
      setLoading(false);
    }
  }, []); // No dependencies — loaderRef is stable

  React.useEffect(() => {
    let disposed = false;

    const load = async (): Promise<void> => {
      if (disposed) {
        return;
      }
      setLoading(true);
      await refresh();
    };

    load().catch(() => undefined);

    const interval = intervalMs > 0
      ? window.setInterval(() => {
        refresh().catch(() => undefined);
      }, intervalMs)
      : undefined;

    const handleFocus = (): void => {
      refresh().catch(() => undefined);
    };

    const handleVisible = (): void => {
      if (document.visibilityState === 'visible') {
        refresh().catch(() => undefined);
      }
    };

    if (refreshOnFocus) {
      window.addEventListener('focus', handleFocus);
    }
    if (refreshOnVisible) {
      document.addEventListener('visibilitychange', handleVisible);
    }

    return () => {
      disposed = true;
      if (interval !== undefined) {
        window.clearInterval(interval);
      }
      if (refreshOnFocus) {
        window.removeEventListener('focus', handleFocus);
      }
      if (refreshOnVisible) {
        document.removeEventListener('visibilitychange', handleVisible);
      }
    };
  }, [intervalMs, refreshOnFocus, refreshOnVisible, refresh]);

  return { data, loading, error, refresh };
}
