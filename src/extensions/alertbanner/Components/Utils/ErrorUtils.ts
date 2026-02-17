import { logger } from '../Services/LoggerService';

export interface IErrorInfo {
  message: string;
  code?: string;
  status?: number;
  originalError: any;
}

export class ErrorUtils {
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

  public static isRetryableError(error: any): boolean {
    if (!error) {
      return false;
    }

    const retryableStatusCodes = [
      408, // Request Timeout
      429, // Too Many Requests
      500, // Internal Server Error
      502, // Bad Gateway
      503, // Service Unavailable
      504  // Gateway Timeout
    ];

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

    if (status && retryableStatusCodes.includes(status)) {
      return true;
    }

    return retryableMessages.some(msg => message.includes(msg));
  }

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

  public static getErrorMessage(error: any): string {
    if (!error) {
      return 'Unknown error';
    }

    if (typeof error === 'string') {
      return error;
    }

    if (error instanceof Error) {
      return error.message;
    }

    if (error.message) {
      return String(error.message);
    }

    if (error.error) {
      return this.getErrorMessage(error.error);
    }

    if (error.statusText) {
      return String(error.statusText);
    }

    try {
      return String(error);
    } catch {
      return 'Unknown error';
    }
  }

  public static getErrorStatus(error: any): number | null {
    if (!error) {
      return null;
    }

    if (typeof error.status === 'number') {
      return error.status;
    }

    if (typeof error.statusCode === 'number') {
      return error.statusCode;
    }

    if (error.response && typeof error.response.status === 'number') {
      return error.response.status;
    }

    const message = this.getErrorMessage(error);
    const statusMatch = message.match(/\b(4\d{2}|5\d{2})\b/);
    if (statusMatch) {
      return parseInt(statusMatch[1], 10);
    }

    return null;
  }

  public static getErrorCode(error: any): string | null {
    if (!error) {
      return null;
    }

    if (error.code && typeof error.code === 'string') {
      return error.code;
    }

    if (error.response && error.response.code) {
      return String(error.response.code);
    }

    return null;
  }

  public static getErrorInfo(error: any): IErrorInfo {
    return {
      message: this.getErrorMessage(error),
      code: this.getErrorCode(error) || undefined,
      status: this.getErrorStatus(error) || undefined,
      originalError: error
    };
  }

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
      const message = this.getErrorMessage(error);
      return message || 'Invalid input. Please check your data and try again.';
    }

    return defaultMessage;
  }

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

  public static shouldShowToUser(error: any): boolean {
    if (!error) {
      return false;
    }

    if (this.isRetryableError(error) && !this.isAccessDeniedError(error)) {
      return false;
    }

    return true;
  }

  public static toError(error: any): Error {
    if (error instanceof Error) {
      return error;
    }

    const message = this.getErrorMessage(error);
    const err = new Error(message);

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
