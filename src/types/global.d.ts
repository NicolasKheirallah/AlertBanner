/**
 * Global type definitions for Alert Banner extension
 */

declare global {
  interface Window {
    // Alert Banner debug flag
    __ALERT_BANNER_DEBUG?: boolean;

    // Application Insights (if using)
    appInsights?: {
      trackException: (exception: { exception: Error; properties?: Record<string, any> }) => void;
      trackEvent: (event: { name: string; properties?: Record<string, any> }) => void;
    };

    // SharePoint specific globals
    _spPageContextInfo?: {
      webAbsoluteUrl: string;
      siteAbsoluteUrl: string;
      userId: number;
      userDisplayName: string;
      userLoginName: string;
    };
  }
}

// SharePoint specific type augmentations
declare module "@microsoft/sp-http" {
  interface MSGraphClientV3 {
    api(path: string): MSGraphRequest;
  }

  interface MSGraphRequest {
    get(): Promise<any>;
    post(data: any): Promise<any>;
    patch(data: any): Promise<any>;
    delete(): Promise<any>;
    version(version: string): MSGraphRequest;
    option(key: string, value: any): MSGraphRequest;
    options(options: { [key: string]: any }): MSGraphRequest;
    select(properties: string): MSGraphRequest;
    expand(properties: string): MSGraphRequest;
    filter(filter: string): MSGraphRequest;
    orderby(orderBy: string): MSGraphRequest;
    top(count: number): MSGraphRequest;
    skip(count: number): MSGraphRequest;
    header(name: string, value: string): MSGraphRequest;
  }
}

// Utility types for better type safety
export type SafeString = string & { __brand: 'SafeString' };
export type Url = string & { __brand: 'Url' };
export type Guid = string & { __brand: 'Guid' };
export type ISODateString = string & { __brand: 'ISODateString' };

// Generic result types
export type Result<TSuccess, TError = Error> =
  | { success: true; data: TSuccess; error?: never }
  | { success: false; error: TError; data?: never };

export type AsyncResult<TSuccess, TError = Error> = Promise<Result<TSuccess, TError>>;

// Form validation types
export type ValidationState = 'valid' | 'invalid' | 'pending' | 'unknown';

export interface ValidationResult<T = any> {
  state: ValidationState;
  errors: string[];
  warnings: string[];
  value?: T;
}

// Export empty object to make this a module
export {};
