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

export class RetryUtils {
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

        const delay = this.calculateDelay(attempt, baseDelay, maxDelay, useExponentialBackoff, useJitter);

        logger.warn(
          'RetryUtils',
          `Attempt ${attempt}/${maxRetries} failed, retrying in ${delay}ms`,
          ErrorUtils.getErrorInfo(error)
        );

        if (onRetry) {
          onRetry(error, attempt, delay);
        }

        await this.delay(delay);
      }
    }

    throw ErrorUtils.toError(lastError || new Error('Maximum retry attempts exceeded'));
  }

  public static calculateDelay(
    attempt: number,
    baseDelay: number,
    maxDelay: number,
    useExponentialBackoff: boolean,
    useJitter: boolean
  ): number {
    let delay: number;

    if (useExponentialBackoff) {
      delay = baseDelay * Math.pow(2, attempt - 1);
    } else {
      delay = baseDelay * attempt;
    }

    if (useJitter) {
      const jitterAmount = delay * 0.3; // 30% jitter
      delay = delay + (Math.random() * jitterAmount * 2 - jitterAmount);
    }

    return Math.min(delay, maxDelay);
  }

  public static delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
