export enum LogLevel {
  DEBUG = 0,
  INFO = 1,
  WARN = 2,
  ERROR = 3,
  FATAL = 4,
}

export interface ILogEntry {
  timestamp: string;
  level: LogLevel;
  component: string;
  message: string;
  data?: unknown;
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
  metadata?: unknown;
}

export class LoggerService {
  private static _instance: LoggerService;
  private logLevel: LogLevel = LogLevel.INFO;
  private maxLogEntries: number = 1000;
  private logEntries: ILogEntry[] = [];
  private sessionId: string;
  private buildVersion: string = "2.0.0";
  private isDevelopment: boolean;
  private performanceMetrics: IPerformanceMetric[] = [];

  private constructor() {
    this.sessionId = this.generateSessionId();
    this.isDevelopment = this.detectDevelopmentMode();
    this.logLevel = this.isDevelopment ? LogLevel.DEBUG : LogLevel.INFO;

    this.setupGlobalErrorHandling();
    this.setupLogCleanup();
  }

  public static getInstance(): LoggerService {
    if (!LoggerService._instance) {
      LoggerService._instance = new LoggerService();
    }
    return LoggerService._instance;
  }

  private generateSessionId(): string {
    return `${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
  }

  private detectDevelopmentMode(): boolean {
    if (
      typeof process !== "undefined" &&
      typeof process.env !== "undefined" &&
      process.env.NODE_ENV === "development"
    ) {
      return true;
    }

    try {
      const hostname = window.location?.hostname || "";
      const queryString =
        window.location?.search || document.location?.search || "";
      const explicitDebugFlag = window.__ALERT_BANNER_DEBUG === true;

      if (explicitDebugFlag) {
        return true;
      }

      const isLocalHost =
        hostname.includes("localhost") || hostname.includes("127.0.0.1");
      const debugQuery = queryString.includes("debug=true");

      return isLocalHost || debugQuery;
    } catch {
      return false;
    }
  }

  private setupGlobalErrorHandling(): void {
    window.addEventListener("unhandledrejection", (event) => {
      const stack = event.reason?.stack || "";
      const isOurCode =
        stack.includes("alert-banner") || stack.includes("AlertBanner");

      if (isOurCode) {
        this.error("GlobalError", "Unhandled promise rejection", {
          reason: event.reason,
          promise: event.promise?.toString(),
        });
        event.preventDefault();
      }
    });

    window.addEventListener("error", (event) => {
      const filename = event.filename || "";
      const isOurCode =
        filename.includes("alert-banner") || filename.includes("AlertBanner");

      if (isOurCode) {
        this.error("GlobalError", "Uncaught error", {
          message: event.message,
          filename: event.filename,
          lineno: event.lineno,
          colno: event.colno,
          error: event.error,
        });
      }
    });
  }

  private setupLogCleanup(): void {
    setInterval(
      () => {
        if (this.logEntries.length > this.maxLogEntries) {
          this.logEntries = this.logEntries.slice(-this.maxLogEntries);
        }
      },
      5 * 60 * 1000,
    );
  }

  private createLogEntry(
    level: LogLevel,
    component: string,
    message: string,
    data?: unknown,
    error?: Error,
  ): ILogEntry {
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
      correlationId: this.generateCorrelationId(),
    };
  }

  private generateCorrelationId(): string {
    return `${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
  }

  private shouldLog(level: LogLevel): boolean {
    return level >= this.logLevel;
  }

  private writeLog(entry: ILogEntry): void {
    this.logEntries.push(entry);

    const consoleMethod = this.getConsoleMethod(entry.level);
    const prefix = `[${entry.component}]`;

    if (entry.error) {
      consoleMethod(prefix, entry.message, entry.error, entry.data || "");
    } else if (entry.data) {
      consoleMethod(prefix, entry.message, entry.data);
    } else {
      consoleMethod(prefix, entry.message);
    }

    if (!this.isDevelopment && entry.level >= LogLevel.ERROR) {
      this.sendToExternalLogging(entry);
    }
  }

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

  private sendToExternalLogging(entry: ILogEntry): void {
    try {
      if (window.appInsights) {
        window.appInsights.trackException({
          exception: entry.error || new Error(entry.message),
          properties: {
            component: entry.component,
            sessionId: entry.sessionId,
            correlationId: entry.correlationId,
            ...(typeof entry.data === "object" && entry.data !== null
              ? entry.data
              : { data: entry.data }),
          },
        });
      }
    } catch (error) {
      this.getConsoleMethod(LogLevel.ERROR)(
        "Failed to send log to external service:",
        error,
      );
    }
  }

  public debug(component: string, message: string, data?: unknown): void {
    if (this.shouldLog(LogLevel.DEBUG)) {
      const entry = this.createLogEntry(
        LogLevel.DEBUG,
        component,
        message,
        data,
      );
      this.writeLog(entry);
    }
  }

  public info(component: string, message: string, data?: unknown): void {
    if (this.shouldLog(LogLevel.INFO)) {
      const entry = this.createLogEntry(
        LogLevel.INFO,
        component,
        message,
        data,
      );
      this.writeLog(entry);
    }
  }

  public warn(component: string, message: string, data?: unknown): void {
    if (this.shouldLog(LogLevel.WARN)) {
      const entry = this.createLogEntry(
        LogLevel.WARN,
        component,
        message,
        data,
      );
      this.writeLog(entry);
    }
  }

  public error(
    component: string,
    message: string,
    error?: Error | unknown,
    data?: unknown,
  ): void {
    if (this.shouldLog(LogLevel.ERROR)) {
      const errorObj =
        error instanceof Error ? error : new Error(String(error));
      const entry = this.createLogEntry(
        LogLevel.ERROR,
        component,
        message,
        data,
        errorObj,
      );
      this.writeLog(entry);
    }
  }

  public fatal(
    component: string,
    message: string,
    error?: Error | unknown,
    data?: unknown,
  ): void {
    if (this.shouldLog(LogLevel.FATAL)) {
      const errorObj =
        error instanceof Error ? error : new Error(String(error));
      const entry = this.createLogEntry(
        LogLevel.FATAL,
        component,
        message,
        data,
        errorObj,
      );
      this.writeLog(entry);
    }
  }

  public logApiCall(
    component: string,
    method: string,
    url: string,
    status?: number,
    duration?: number,
    error?: Error,
  ): void {
    const logData = {
      method,
      url,
      status,
      duration: duration ? `${duration}ms` : undefined,
      timestamp: new Date().toISOString(),
    };

    if (error || (status && status >= 400)) {
      this.error(
        component,
        `API call failed: ${method} ${url}`,
        error,
        logData,
      );
    } else {
      this.info(component, `API call successful: ${method} ${url}`, logData);
    }
  }

  public startPerformanceTracking(name: string): () => void {
    const startTime = performance.now();

    return () => {
      const duration = performance.now() - startTime;
      const metric: IPerformanceMetric = {
        name,
        duration,
        timestamp: new Date().toISOString(),
      };

      this.performanceMetrics.push(metric);
      this.debug(
        "Performance",
        `${name} completed in ${duration.toFixed(2)}ms`,
      );

      if (this.performanceMetrics.length > 100) {
        this.performanceMetrics = this.performanceMetrics.slice(-100);
      }
    };
  }

  public logUserAction(
    component: string,
    action: string,
    metadata?: unknown,
  ): void {
    this.info(component, `User action: ${action}`, {
      action,
      metadata,
      timestamp: new Date().toISOString(),
    });
  }
}

export const logger = LoggerService.getInstance();
