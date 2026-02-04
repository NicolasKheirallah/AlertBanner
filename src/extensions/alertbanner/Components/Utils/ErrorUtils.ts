/**
 * Error detection and handling utilities
 * Consolidates duplicate error checking patterns from across the codebase
 */

import { logger } from '../Services/LoggerService';

export interface IErrorInfo {
  message: string;
  code?: string;
  status?: number;
  originalError: any;
}

/**
 * Utility class for error detection and handling
 */
export class ErrorUtils {
  /**
   * Check if error is an access denied error (403 or permission-related)
   */
  public static isAccessDeniedError(error: any): boolean {
    if (!error) {
      return false;
    }

    const message = this.getErrorMessage(error).toLowerCase();
    const status = this.getErrorStatus(error);

    return (
      message.includes('access denied') ||
      message.includes('permission') ||
      message.includes('unauthorized') ||
      message.includes('forbidden') ||
      status === 403
    );
  }

  /**
   * Check if error is a list/resource not found error (404)
   */
  public static isListNotFoundError(error: any): boolean {
    if (!error) {
      return false;
    }

    const message = this.getErrorMessage(error).toLowerCase();
    const status = this.getErrorStatus(error);
    const code = (this.getErrorCode(error) || "").toLowerCase();

    return (
      code === 'itemnotfound' ||
      code === 'notfound' ||
      message.includes('list_not_found') ||
      message.includes('not found') ||
      message.includes('does not exist') ||
      status === 404
    );
  }

  /**
   * Check if error is retryable (transient network/server errors)
   */
  public static isRetryableError(error: any): boolean {
    if (!error) {
      return false;
    }

    // HTTP status codes that are typically retryable
    const retryableStatusCodes = [
      408, // Request Timeout
      429, // Too Many Requests (rate limiting)
      500, // Internal Server Error
      502, // Bad Gateway
      503, // Service Unavailable
      504  // Gateway Timeout
    ];

    // Error message patterns that indicate transient issues
    const retryableMessages = [
      'timeout',
      'network',
      'throttled',
      'temporarily unavailable',
      'service unavailable',
      'connection',
      'econnreset',
      'etimedout',
      'socket hang up'
    ];

    const message = this.getErrorMessage(error).toLowerCase();
    const status = this.getErrorStatus(error);

    // Check if status code is retryable
    if (status && retryableStatusCodes.includes(status)) {
      return true;
    }

    // Check if error message indicates a retryable error
    return retryableMessages.some(msg => message.includes(msg));
  }

  /**
   * Check if error is a network error
   */
  public static isNetworkError(error: any): boolean {
    if (!error) {
      return false;
    }

    const message = this.getErrorMessage(error).toLowerCase();

    return (
      message.includes('network') ||
      message.includes('connection') ||
      message.includes('offline') ||
      message.includes('econnrefused') ||
      message.includes('enotfound') ||
      message.includes('etimedout') ||
      message.includes('econnreset')
    );
  }

  /**
   * Check if error is a validation error
   */
  public static isValidationError(error: any): boolean {
    if (!error) {
      return false;
    }

    const message = this.getErrorMessage(error).toLowerCase();
    const status = this.getErrorStatus(error);

    return (
      message.includes('validation') ||
      message.includes('invalid') ||
      message.includes('required') ||
      status === 400
    );
  }

  /**
   * Check if error is an authentication error
   */
  public static isAuthenticationError(error: any): boolean {
    if (!error) {
      return false;
    }

    const message = this.getErrorMessage(error).toLowerCase();
    const status = this.getErrorStatus(error);

    return (
      message.includes('authentication') ||
      message.includes('unauthorized') ||
      message.includes('not authenticated') ||
      status === 401
    );
  }

  /**
   * Extract error message from various error types
   */
  public static getErrorMessage(error: any): string {
    if (!error) {
      return 'Unknown error';
    }

    // Handle string errors
    if (typeof error === 'string') {
      return error;
    }

    // Handle Error objects
    if (error instanceof Error) {
      return error.message;
    }

    // Handle objects with message property
    if (error.message) {
      return String(error.message);
    }

    // Handle objects with error property
    if (error.error) {
      return this.getErrorMessage(error.error);
    }

    // Handle response errors (like from fetch)
    if (error.statusText) {
      return String(error.statusText);
    }

    // Fallback to stringification
    try {
      return String(error);
    } catch {
      return 'Unknown error';
    }
  }

  /**
   * Extract HTTP status code from error
   */
  public static getErrorStatus(error: any): number | null {
    if (!error) {
      return null;
    }

    // Direct status property
    if (typeof error.status === 'number') {
      return error.status;
    }

    // statusCode property (common in Node.js errors)
    if (typeof error.statusCode === 'number') {
      return error.statusCode;
    }

    // Response object
    if (error.response && typeof error.response.status === 'number') {
      return error.response.status;
    }

    // Try to extract from error message
    const message = this.getErrorMessage(error);
    const statusMatch = message.match(/\b(4\d{2}|5\d{2})\b/);
    if (statusMatch) {
      return parseInt(statusMatch[1], 10);
    }

    return null;
  }

  /**
   * Extract error code from error
   */
  public static getErrorCode(error: any): string | null {
    if (!error) {
      return null;
    }

    // Direct code property
    if (error.code && typeof error.code === 'string') {
      return error.code;
    }

    // Error response code
    if (error.response && error.response.code) {
      return String(error.response.code);
    }

    return null;
  }

  /**
   * Create structured error info from any error type
   */
  public static getErrorInfo(error: any): IErrorInfo {
    return {
      message: this.getErrorMessage(error),
      code: this.getErrorCode(error) || undefined,
      status: this.getErrorStatus(error) || undefined,
      originalError: error
    };
  }

  /**
   * Log error with appropriate level based on error type
   */
  public static logError(context: string, error: any, additionalData?: any): void {
    const errorInfo = this.getErrorInfo(error);

    if (this.isRetryableError(error) || this.isNetworkError(error)) {
      logger.warn(context, `Transient error: ${errorInfo.message}`, {
        ...errorInfo,
        ...additionalData
      });
    } else {
      logger.error(context, errorInfo.message, {
        ...errorInfo,
        ...additionalData
      });
    }
  }

  /**
   * Create user-friendly error message
   */
  public static getUserFriendlyMessage(error: any, defaultMessage: string = 'An unexpected error occurred'): string {
    if (!error) {
      return defaultMessage;
    }

    if (this.isNetworkError(error)) {
      return 'Network connection issue. Please check your internet connection and try again.';
    }

    if (this.isAccessDeniedError(error)) {
      return 'You do not have permission to perform this action.';
    }

    if (this.isAuthenticationError(error)) {
      return 'Authentication required. Please sign in and try again.';
    }

    if (this.isListNotFoundError(error)) {
      return 'The requested resource was not found.';
    }

    if (this.isValidationError(error)) {
      // For validation errors, try to return the actual message as it's usually user-friendly
      const message = this.getErrorMessage(error);
      return message || 'Invalid input. Please check your data and try again.';
    }

    // For unknown errors, use default message
    return defaultMessage;
  }

  /**
   * Wrap async operation with error handling
   */
  public static async tryExecute<T>(
    operation: () => Promise<T>,
    context: string,
    options: {
      onError?: (error: any) => void;
      defaultValue?: T;
      logError?: boolean;
    } = {}
  ): Promise<T | null> {
    try {
      return await operation();
    } catch (error) {
      if (options.logError !== false) {
        this.logError(context, error);
      }

      if (options.onError) {
        options.onError(error);
      }

      return options.defaultValue ?? null;
    }
  }

  /**
   * Check if error should be shown to user
   */
  public static shouldShowToUser(error: any): boolean {
    if (!error) {
      return false;
    }

    // Don't show transient errors that might auto-resolve
    if (this.isRetryableError(error) && !this.isAccessDeniedError(error)) {
      return false;
    }

    return true;
  }

  /**
   * Create Error object from any error type
   */
  public static toError(error: any): Error {
    if (error instanceof Error) {
      return error;
    }

    const message = this.getErrorMessage(error);
    const err = new Error(message);

    // Preserve status and code if available
    const status = this.getErrorStatus(error);
    const code = this.getErrorCode(error);

    if (status) {
      (err as any).status = status;
    }

    if (code) {
      (err as any).code = code;
    }

    return err;
  }
}
