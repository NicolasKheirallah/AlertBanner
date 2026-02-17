import { useState, useCallback, useRef, useEffect } from 'react';
import { logger } from '../Services/LoggerService';
import { ErrorUtils } from '../Utils/ErrorUtils';

export interface IAsyncOperationOptions<T> {
  onSuccess?: (data: T) => void;
  onError?: (error: Error) => void;
  successMessage?: string;
  errorMessage?: string;
  logErrors?: boolean;
  resetOnUnmount?: boolean;
}

export interface IAsyncOperationState<T> {
  loading: boolean;
  error: Error | null;
  data: T | null;
  message: { type: 'success' | 'error' | 'info' | 'warning'; text: string } | null;
}

export interface IAsyncOperationReturn<T, TArgs extends any[]> {
  loading: boolean;
  error: Error | null;
  data: T | null;
  message: { type: 'success' | 'error' | 'info' | 'warning'; text: string } | null;
  execute: (...args: TArgs) => Promise<T | null>;
  reset: () => void;
  setMessage: (message: { type: 'success' | 'error' | 'info' | 'warning'; text: string } | null) => void;
  clearError: () => void;
}

export function useAsyncOperation<T, TArgs extends any[] = []>(
  operation: (...args: TArgs) => Promise<T>,
  options: IAsyncOperationOptions<T> = {}
): IAsyncOperationReturn<T, TArgs> {
  const [state, setState] = useState<IAsyncOperationState<T>>({
    loading: false,
    error: null,
    data: null,
    message: null
  });

  const isMountedRef = useRef(true);
  const abortControllerRef = useRef<AbortController | null>(null);

  // Track mount status
  useEffect(() => {
    isMountedRef.current = true;

    return () => {
      isMountedRef.current = false;
      // Abort any pending operations on unmount
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
    };
  }, []);

  // Reset state on unmount if requested
  useEffect(() => {
    return () => {
      if (options.resetOnUnmount && isMountedRef.current) {
        setState({
          loading: false,
          error: null,
          data: null,
          message: null
        });
      }
    };
  }, [options.resetOnUnmount]);

  const execute = useCallback(
    async (...args: TArgs): Promise<T | null> => {
      // Abort any previous operation
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }

      // Create new abort controller for this operation
      abortControllerRef.current = new AbortController();

      setState({
        loading: true,
        error: null,
        data: null,
        message: null
      });

      try {
        const result = await operation(...args);

        // Only update state if component is still mounted
        if (isMountedRef.current) {
          setState({
            loading: false,
            error: null,
            data: result,
            message: options.successMessage
              ? { type: 'success', text: options.successMessage }
              : null
          });

          if (options.onSuccess) {
            options.onSuccess(result);
          }
        }

        return result;
      } catch (err: any) {
        const error = ErrorUtils.toError(err);

        // Only update state if component is still mounted
        if (isMountedRef.current) {
          const errorMessage = options.errorMessage
            ? options.errorMessage
            : ErrorUtils.getUserFriendlyMessage(error);

          setState({
            loading: false,
            error,
            data: null,
            message: { type: 'error', text: errorMessage }
          });

          if (options.logErrors !== false) {
            logger.error('useAsyncOperation', 'Async operation failed', {
              error,
              errorMessage
            });
          }

          if (options.onError) {
            options.onError(error);
          }
        }

        return null;
      }
    },
    [operation, options]
  );

  const reset = useCallback(() => {
    setState({
      loading: false,
      error: null,
      data: null,
      message: null
    });
  }, []);

  const setMessage = useCallback(
    (message: { type: 'success' | 'error' | 'info' | 'warning'; text: string } | null) => {
      setState((prev) => ({
        ...prev,
        message
      }));
    },
    []
  );

  const clearError = useCallback(() => {
    setState((prev) => ({
      ...prev,
      error: null,
      message: prev.message?.type === 'error' ? null : prev.message
    }));
  }, []);

  return {
    loading: state.loading,
    error: state.error,
    data: state.data,
    message: state.message,
    execute,
    reset,
    setMessage,
    clearError
  };
}

export function useAsyncOperationImmediate<T, TArgs extends any[] = []>(
  operation: (...args: TArgs) => Promise<T>,
  options: IAsyncOperationOptions<T> = {},
  dependencies: any[] = []
): IAsyncOperationReturn<T, TArgs> & { refresh: () => Promise<T | null> } {
  const asyncOp = useAsyncOperation(operation, options);

  // Execute on mount and when dependencies change
  useEffect(() => {
    asyncOp.execute(...([] as unknown as TArgs));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, dependencies);

  const refresh = useCallback(() => {
    return asyncOp.execute(...([] as unknown as TArgs));
  }, [asyncOp]);

  return {
    ...asyncOp,
    refresh
  };
}

export function useAsyncOperationPolling<T, TArgs extends any[] = []>(
  operation: (...args: TArgs) => Promise<T>,
  options: IAsyncOperationOptions<T> & {
    interval: number;
    startImmediately?: boolean;
  }
): IAsyncOperationReturn<T, TArgs> & {
  startPolling: (...args: TArgs) => void;
  stopPolling: () => void;
  isPolling: boolean;
} {
  const asyncOp = useAsyncOperation(operation, options);
  const [isPolling, setIsPolling] = useState(false);
  const intervalRef = useRef<any>(null);
  const argsRef = useRef<TArgs | null>(null);

  const stopPolling = useCallback(() => {
    if (intervalRef.current) {
      clearInterval(intervalRef.current);
      intervalRef.current = null;
    }
    setIsPolling(false);
  }, []);

  const startPolling = useCallback(
    (...args: TArgs) => {
      // Stop any existing polling
      stopPolling();

      // Store args for polling
      argsRef.current = args;
      setIsPolling(true);

      // Execute immediately
      asyncOp.execute(...args);

      // Start interval
      intervalRef.current = setInterval(() => {
        if (argsRef.current) {
          asyncOp.execute(...argsRef.current);
        }
      }, options.interval);
    },
    [asyncOp, options.interval, stopPolling]
  );

  // Auto-start if requested
  useEffect(() => {
    if (options.startImmediately) {
      startPolling(...([] as unknown as TArgs));
    }

    // Cleanup on unmount
    return () => {
      stopPolling();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return {
    ...asyncOp,
    startPolling,
    stopPolling,
    isPolling
  };
}
