/**
 * Production-ready logging service for Alert Banner extension
 * Provides structured logging, error tracking, and performance monitoring
 */

export enum LogLevel {
  DEBUG = 0,
  INFO = 1,
  WARN = 2,
  ERROR = 3,
  FATAL = 4
}

export interface ILogEntry {
  timestamp: string;
  level: LogLevel;
  component: string;
  message: string;
  data?: any;
  error?: Error;
  userId?: string;
  sessionId: string;
  buildVersion: string;
  userAgent: string;
  url: string;
  correlationId?: string;
}

export interface IPerformanceMetric {
  name: string;
  duration: number;
  timestamp: string;
  metadata?: any;
}

export class LoggerService {
  private static _instance: LoggerService;
  private logLevel: LogLevel = LogLevel.INFO; // Default to INFO in production
  private maxLogEntries: number = 1000;
  private logEntries: ILogEntry[] = [];
  private sessionId: string;
  private buildVersion: string = '2.0.0';
  private isDevelopment: boolean;
  private performanceMetrics: IPerformanceMetric[] = [];

  private constructor() {
    this.sessionId = this.generateSessionId();
    this.isDevelopment = this.detectDevelopmentMode();
    this.logLevel = this.isDevelopment ? LogLevel.DEBUG : LogLevel.INFO;
    
    // Listen for unhandled errors
    this.setupGlobalErrorHandling();
    
    // Periodic cleanup of old logs
    this.setupLogCleanup();
  }

  public static getInstance(): LoggerService {
    if (!LoggerService._instance) {
      LoggerService._instance = new LoggerService();
    }
    return LoggerService._instance;
  }

  /**
   * Generate unique session ID
   */
  private generateSessionId(): string {
    return `${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
  }

  /**
   * Detect if running in development mode
   */
  private detectDevelopmentMode(): boolean {
    if (typeof process !== 'undefined' && typeof process.env !== 'undefined' && process.env.NODE_ENV === 'development') {
      return true;
    }

    try {
      const hostname = window.location?.hostname || '';
      const queryString = window.location?.search || document.location?.search || '';
      const explicitDebugFlag = (window as any).__ALERT_BANNER_DEBUG === true;

      if (explicitDebugFlag) {
        return true;
      }

      const isLocalHost = hostname.includes('localhost') || hostname.includes('127.0.0.1');
      const debugQuery = queryString.includes('debug=true');

      return isLocalHost || debugQuery;
    } catch {
      return false;
    }
  }

  /**
   * Setup global error handling
   */
  private setupGlobalErrorHandling(): void {
    // Handle unhandled promise rejections - only log if from our code
    window.addEventListener('unhandledrejection', (event) => {
      // Check if error is from our Alert Banner code
      const stack = event.reason?.stack || '';
      const isOurCode = stack.includes('alert-banner') || stack.includes('AlertBanner');

      if (isOurCode) {
        this.error('GlobalError', 'Unhandled promise rejection', {
          reason: event.reason,
          promise: event.promise?.toString()
        });
        // Prevent default to avoid duplicate console errors
        event.preventDefault();
      }
    });

    // Handle uncaught errors - only log if from our code
    window.addEventListener('error', (event) => {
      // Check if error is from our Alert Banner code
      const filename = event.filename || '';
      const isOurCode = filename.includes('alert-banner') || filename.includes('AlertBanner');

      if (isOurCode) {
        this.error('GlobalError', 'Uncaught error', {
          message: event.message,
          filename: event.filename,
          lineno: event.lineno,
          colno: event.colno,
          error: event.error
        });
      }
    });
  }

  /**
   * Setup periodic log cleanup
   */
  private setupLogCleanup(): void {
    // Clean up old logs every 5 minutes
    setInterval(() => {
      if (this.logEntries.length > this.maxLogEntries) {
        this.logEntries = this.logEntries.slice(-this.maxLogEntries);
      }
    }, 5 * 60 * 1000);
  }

  /**
   * Create log entry
   */
  private createLogEntry(level: LogLevel, component: string, message: string, data?: any, error?: Error): ILogEntry {
    return {
      timestamp: new Date().toISOString(),
      level,
      component,
      message,
      data,
      error,
      sessionId: this.sessionId,
      buildVersion: this.buildVersion,
      userAgent: navigator.userAgent,
      url: window.location.href,
      correlationId: this.generateCorrelationId()
    };
  }

  /**
   * Generate correlation ID for request tracking
   */
  private generateCorrelationId(): string {
    return `${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
  }

  /**
   * Check if log level should be logged
   */
  private shouldLog(level: LogLevel): boolean {
    return level >= this.logLevel;
  }

  /**
   * Write log to storage and console
   */
  private writeLog(entry: ILogEntry): void {
    // Store in memory
    this.logEntries.push(entry);

    // Console output with appropriate styling
    const consoleMethod = this.getConsoleMethod(entry.level);
    const prefix = `[${entry.component}]`;
    
    if (entry.error) {
      consoleMethod(prefix, entry.message, entry.error, entry.data || '');
    } else if (entry.data) {
      consoleMethod(prefix, entry.message, entry.data);
    } else {
      consoleMethod(prefix, entry.message);
    }

    // In production, consider sending to external logging service
    if (!this.isDevelopment && entry.level >= LogLevel.ERROR) {
      this.sendToExternalLogging(entry);
    }
  }

  /**
   * Get appropriate console method for log level
   */
  private getConsoleMethod(level: LogLevel): typeof console.log {
    switch (level) {
      case LogLevel.DEBUG:
        return console.debug;
      case LogLevel.INFO:
        return console.info;
      case LogLevel.WARN:
        return console.warn;
      case LogLevel.ERROR:
      case LogLevel.FATAL:
        return console.error;
      default:
        return console.log;
    }
  }

  /**
   * Send critical logs to external service
   * Integrate with Application Insights or similar monitoring service
   */
  private sendToExternalLogging(entry: ILogEntry): void {
    try {
      // Integration point for external logging services
      // Configure your preferred service (Application Insights, LogRocket, etc.)
      if ((window as any).appInsights) {
        (window as any).appInsights.trackException({
          exception: entry.error || new Error(entry.message),
          properties: {
            component: entry.component,
            sessionId: entry.sessionId,
            correlationId: entry.correlationId,
            ...entry.data
          }
        });
      }
    } catch (error) {
      // Silently fail external logging to prevent disruption
      this.getConsoleMethod(LogLevel.ERROR)('Failed to send log to external service:', error);
    }
  }

  /**
   * Debug level logging
   */
  public debug(component: string, message: string, data?: any): void {
    if (this.shouldLog(LogLevel.DEBUG)) {
      const entry = this.createLogEntry(LogLevel.DEBUG, component, message, data);
      this.writeLog(entry);
    }
  }

  /**
   * Info level logging
   */
  public info(component: string, message: string, data?: any): void {
    if (this.shouldLog(LogLevel.INFO)) {
      const entry = this.createLogEntry(LogLevel.INFO, component, message, data);
      this.writeLog(entry);
    }
  }

  /**
   * Warning level logging
   */
  public warn(component: string, message: string, data?: any): void {
    if (this.shouldLog(LogLevel.WARN)) {
      const entry = this.createLogEntry(LogLevel.WARN, component, message, data);
      this.writeLog(entry);
    }
  }

  /**
   * Error level logging
   */
  public error(component: string, message: string, error?: Error | any, data?: any): void {
    if (this.shouldLog(LogLevel.ERROR)) {
      const errorObj = error instanceof Error ? error : new Error(String(error));
      const entry = this.createLogEntry(LogLevel.ERROR, component, message, data, errorObj);
      this.writeLog(entry);
    }
  }

  /**
   * Fatal level logging
   */
  public fatal(component: string, message: string, error?: Error | any, data?: any): void {
    if (this.shouldLog(LogLevel.FATAL)) {
      const errorObj = error instanceof Error ? error : new Error(String(error));
      const entry = this.createLogEntry(LogLevel.FATAL, component, message, data, errorObj);
      this.writeLog(entry);
    }
  }

  /**
   * Performance monitoring
   */
  public startPerformanceTracking(name: string): () => void {
    const startTime = performance.now();
    
    return () => {
      const duration = performance.now() - startTime;
      const metric: IPerformanceMetric = {
        name,
        duration,
        timestamp: new Date().toISOString()
      };
      
      this.performanceMetrics.push(metric);
      this.debug('Performance', `${name} completed in ${duration.toFixed(2)}ms`);
      
      // Clean up old metrics
      if (this.performanceMetrics.length > 100) {
        this.performanceMetrics = this.performanceMetrics.slice(-100);
      }
    };
  }

  /**
   * Structured API call logging
   */
  public logApiCall(component: string, method: string, url: string, status?: number, duration?: number, error?: Error): void {
    const logData = {
      method,
      url,
      status,
      duration: duration ? `${duration}ms` : undefined,
      timestamp: new Date().toISOString()
    };

    if (error || (status && status >= 400)) {
      this.error(component, `API call failed: ${method} ${url}`, error, logData);
    } else {
      this.info(component, `API call successful: ${method} ${url}`, logData);
    }
  }

  /**
   * User action logging
   */
  public logUserAction(component: string, action: string, metadata?: any): void {
    this.info(component, `User action: ${action}`, {
      action,
      metadata,
      timestamp: new Date().toISOString()
    });
  }

  /**
   * Get recent logs (for debugging)
   */
  public getRecentLogs(count: number = 50): ILogEntry[] {
    return this.logEntries.slice(-count);
  }

  /**
   * Get logs by level
   */
  public getLogsByLevel(level: LogLevel): ILogEntry[] {
    return this.logEntries.filter(entry => entry.level === level);
  }

  /**
   * Get performance metrics
   */
  public getPerformanceMetrics(): IPerformanceMetric[] {
    return [...this.performanceMetrics];
  }

  /**
   * Export logs for debugging
   */
  public exportLogs(): string {
    return JSON.stringify({
      sessionId: this.sessionId,
      buildVersion: this.buildVersion,
      exportTime: new Date().toISOString(),
      logEntries: this.logEntries,
      performanceMetrics: this.performanceMetrics
    }, null, 2);
  }

  /**
   * Clear all logs
   */
  public clearLogs(): void {
    this.logEntries = [];
    this.performanceMetrics = [];
    this.info('LoggerService', 'Logs cleared');
  }

  /**
   * Set log level
   */
  public setLogLevel(level: LogLevel): void {
    this.logLevel = level;
    this.info('LoggerService', `Log level set to ${LogLevel[level]}`);
  }

  /**
   * Get current configuration
   */
  public getConfiguration(): { sessionId: string; logLevel: string; isDevelopment: boolean; buildVersion: string } {
    return {
      sessionId: this.sessionId,
      logLevel: LogLevel[this.logLevel],
      isDevelopment: this.isDevelopment,
      buildVersion: this.buildVersion
    };
  }
}

// Export singleton instance
export const logger = LoggerService.getInstance();
