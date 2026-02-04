/**
 * Retry utilities for handling transient failures
 * Consolidates duplicate retry logic from across the codebase
 */

import { logger } from '../Services/LoggerService';
import { ErrorUtils } from './ErrorUtils';

export interface IRetryOptions {
  maxRetries?: number;
  baseDelay?: number;
  maxDelay?: number;
  useExponentialBackoff?: boolean;
  useJitter?: boolean;
  shouldRetry?: (error: any, attempt: number) => boolean;
  onRetry?: (error: any, attempt: number, delay: number) => void;
  suppressFailureLog?: (error: any, attempt: number) => boolean;
}

export interface IRetryResult<T> {
  success: boolean;
  data?: T;
  error?: Error;
  attempts: number;
}

/**
 * Utility class for retry operations
 */
export class RetryUtils {
  /**
   * Execute operation with automatic retry on failure
   *
   * @example
   * const result = await RetryUtils.executeWithRetry(
   *   async () => {
   *     return await fetch('/api/data');
   *   },
   *   {
   *     maxRetries: 3,
   *     baseDelay: 1000,
   *     useExponentialBackoff: true
   *   }
   * );
   */
  public static async executeWithRetry<T>(
    operation: () => Promise<T>,
    options: IRetryOptions = {}
  ): Promise<T> {
    const {
      maxRetries = 3,
      baseDelay = 1000,
      maxDelay = 30000,
      useExponentialBackoff = true,
      useJitter = true,
      shouldRetry = (error) => ErrorUtils.isRetryableError(error),
      onRetry,
      suppressFailureLog
    } = options;

    let lastError: any;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await operation();
      } catch (error: any) {
        lastError = error;

        const isRetryable = shouldRetry(error, attempt);
        const isLastAttempt = attempt === maxRetries;

        if (!isRetryable || isLastAttempt) {
          const shouldSuppress = suppressFailureLog ? suppressFailureLog(error, attempt) : false;
          if (!shouldSuppress) {
            logger.error(
              'RetryUtils',
              `Operation failed after ${attempt} attempt(s)`,
              ErrorUtils.getErrorInfo(error)
            );
          }
          throw ErrorUtils.toError(error);
        }

        // Calculate delay with exponential backoff and jitter
        const delay = this.calculateDelay(attempt, baseDelay, maxDelay, useExponentialBackoff, useJitter);

        logger.warn(
          'RetryUtils',
          `Attempt ${attempt}/${maxRetries} failed, retrying in ${delay}ms`,
          ErrorUtils.getErrorInfo(error)
        );

        // Notify about retry if callback provided
        if (onRetry) {
          onRetry(error, attempt, delay);
        }

        // Wait before next attempt
        await this.delay(delay);
      }
    }

    // This should never be reached, but TypeScript needs it
    throw ErrorUtils.toError(lastError || new Error('Maximum retry attempts exceeded'));
  }

  /**
   * Execute operation with retry and return result object instead of throwing
   */
  public static async tryExecuteWithRetry<T>(
    operation: () => Promise<T>,
    options: IRetryOptions = {}
  ): Promise<IRetryResult<T>> {
    const maxRetries = options.maxRetries || 3;
    let attempts = 0;

    try {
      const data = await this.executeWithRetry(operation, options);
      return {
        success: true,
        data,
        attempts: attempts + 1
      };
    } catch (error: any) {
      return {
        success: false,
        error: ErrorUtils.toError(error),
        attempts: maxRetries
      };
    }
  }

  /**
   * Calculate delay for next retry attempt
   */
  public static calculateDelay(
    attempt: number,
    baseDelay: number,
    maxDelay: number,
    useExponentialBackoff: boolean,
    useJitter: boolean
  ): number {
    let delay: number;

    if (useExponentialBackoff) {
      // Exponential backoff: delay = baseDelay * 2^(attempt - 1)
      delay = baseDelay * Math.pow(2, attempt - 1);
    } else {
      // Linear backoff: delay = baseDelay * attempt
      delay = baseDelay * attempt;
    }

    // Add jitter (randomization) to prevent thundering herd
    if (useJitter) {
      const jitterAmount = delay * 0.3; // 30% jitter
      delay = delay + (Math.random() * jitterAmount * 2 - jitterAmount);
    }

    // Cap at maximum delay
    return Math.min(delay, maxDelay);
  }

  /**
   * Calculate exponential backoff delay
   */
  public static calculateExponentialBackoff(
    attempt: number,
    baseDelay: number = 1000,
    maxDelay: number = 30000
  ): number {
    return this.calculateDelay(attempt, baseDelay, maxDelay, true, true);
  }

  /**
   * Calculate linear backoff delay
   */
  public static calculateLinearBackoff(
    attempt: number,
    baseDelay: number = 1000,
    maxDelay: number = 30000
  ): number {
    return this.calculateDelay(attempt, baseDelay, maxDelay, false, true);
  }

  /**
   * Simple delay/sleep function
   */
  public static delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Retry operation with timeout
   */
  public static async executeWithRetryAndTimeout<T>(
    operation: () => Promise<T>,
    timeoutMs: number,
    retryOptions: IRetryOptions = {}
  ): Promise<T> {
    return Promise.race([
      this.executeWithRetry(operation, retryOptions),
      this.createTimeoutPromise<T>(timeoutMs)
    ]);
  }

  /**
   * Create a promise that rejects after timeout
   */
  private static createTimeoutPromise<T>(ms: number): Promise<T> {
    return new Promise((_, reject) => {
      setTimeout(() => {
        reject(new Error(`Operation timed out after ${ms}ms`));
      }, ms);
    });
  }

  /**
   * Retry operation until success or max attempts
   * Returns null on failure instead of throwing
   */
  public static async retryUntilSuccess<T>(
    operation: () => Promise<T>,
    options: IRetryOptions = {}
  ): Promise<T | null> {
    try {
      return await this.executeWithRetry(operation, options);
    } catch (error) {
      logger.error('RetryUtils', 'All retry attempts failed', ErrorUtils.getErrorInfo(error));
      return null;
    }
  }

  /**
   * Retry operation with custom backoff strategy
   */
  public static async executeWithCustomBackoff<T>(
    operation: () => Promise<T>,
    backoffStrategy: (attempt: number) => number,
    maxRetries: number = 3,
    shouldRetry?: (error: any, attempt: number) => boolean
  ): Promise<T> {
    return this.executeWithRetry(operation, {
      maxRetries,
      shouldRetry,
      onRetry: async (error, attempt, _) => {
        const customDelay = backoffStrategy(attempt);
        await this.delay(customDelay);
      }
    });
  }

  /**
   * Create a retry-enabled version of an async function
   */
  public static withRetry<TArgs extends any[], TReturn>(
    fn: (...args: TArgs) => Promise<TReturn>,
    options: IRetryOptions = {}
  ): (...args: TArgs) => Promise<TReturn> {
    return async (...args: TArgs): Promise<TReturn> => {
      return this.executeWithRetry(() => fn(...args), options);
    };
  }

  /**
   * Retry with circuit breaker pattern
   * If failures exceed threshold, circuit opens and operations fail fast
   */
  public static createCircuitBreaker<T>(
    operation: () => Promise<T>,
    options: {
      failureThreshold?: number;
      resetTimeout?: number;
      retryOptions?: IRetryOptions;
    } = {}
  ): () => Promise<T> {
    const {
      failureThreshold = 5,
      resetTimeout = 60000, // 1 minute
      retryOptions = {}
    } = options;

    let failureCount = 0;
    let circuitOpen = false;
    let lastFailureTime: number | null = null;

    return async (): Promise<T> => {
      // Check if circuit should be reset
      if (circuitOpen && lastFailureTime) {
        const timeSinceLastFailure = Date.now() - lastFailureTime;
        if (timeSinceLastFailure >= resetTimeout) {
          logger.info('RetryUtils', 'Circuit breaker reset, attempting operation');
          circuitOpen = false;
          failureCount = 0;
        }
      }

      // If circuit is open, fail fast
      if (circuitOpen) {
        throw new Error('Circuit breaker is open - too many failures');
      }

      try {
        const result = await this.executeWithRetry(operation, retryOptions);

        // Reset failure count on success
        failureCount = 0;
        return result;
      } catch (error) {
        failureCount++;
        lastFailureTime = Date.now();

        // Open circuit if threshold exceeded
        if (failureCount >= failureThreshold) {
          circuitOpen = true;
          logger.error(
            'RetryUtils',
            `Circuit breaker opened after ${failureCount} failures`,
            { resetTimeout }
          );
        }

        throw error;
      }
    };
  }

  /**
   * Batch retry operations with rate limiting
   */
  public static async executeBatchWithRetry<T>(
    operations: Array<() => Promise<T>>,
    options: {
      concurrency?: number;
      retryOptions?: IRetryOptions;
      stopOnFirstError?: boolean;
    } = {}
  ): Promise<Array<T | null>> {
    const {
      concurrency = 5,
      retryOptions = {},
      stopOnFirstError = false
    } = options;

    const results: Array<T | null> = [];
    const executing: Array<Promise<void>> = [];

    for (let i = 0; i < operations.length; i++) {
      const operation = operations[i];

      const promise = (async () => {
        try {
          const result = await this.executeWithRetry(operation, retryOptions);
          results[i] = result;
        } catch (error) {
          results[i] = null;
          if (stopOnFirstError) {
            throw error;
          }
        }
      })();

      executing.push(promise);

      // Limit concurrency
      if (executing.length >= concurrency) {
        await Promise.race(executing);
        const completed = executing.findIndex(p => {
          return Promise.race([p, Promise.resolve()]).then(() => true);
        });
        if (completed !== -1) {
          executing.splice(completed, 1);
        }
      }
    }

    // Wait for all remaining operations
    await Promise.all(executing);

    return results;
  }
}
