import { IAlertItem } from "../Alerts/IAlerts";
import { JsonUtils } from "../Utils/JsonUtils";
import { logger } from "./LoggerService";

export interface IStorageOptions {
  expirationTime?: number; // In milliseconds
  userSpecific?: boolean; // Whether to prefix with user ID
}

export class StorageService {
  private static instance: StorageService;
  private userId: string | null = null;
  private defaultExpirationTime = 24 * 60 * 60 * 1000; // 24 hours in milliseconds

  private constructor() {
  }

  public static getInstance(): StorageService {
    if (!StorageService.instance) {
      StorageService.instance = new StorageService();
    }
    return StorageService.instance;
  }

  public setUserId(userId: string): void {
    this.userId = userId;
  }

  public saveToLocalStorage<T>(
    key: string,
    data: T,
    options?: IStorageOptions,
  ): void {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      const storageData = {
        data,
        timestamp: Date.now(),
        expiration: options?.expirationTime || this.defaultExpirationTime,
      };

      const serialized = JsonUtils.safeStringify(storageData);
      if (serialized) {
        localStorage.setItem(fullKey, serialized);
      }
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to save to localStorage: ${key}`,
        error,
      );
    }
  }

  public getFromLocalStorage<T>(
    key: string,
    options?: IStorageOptions,
  ): T | null {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      const data = localStorage.getItem(fullKey);

      if (!data) return null;

      const parsedData = JsonUtils.safeParse(data);
      if (!parsedData) return null;

      if (this.isDataExpired(parsedData)) {
        this.removeFromLocalStorage(key, options);
        return null;
      }

      return parsedData.data as T;
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to read from localStorage: ${key}`,
        error,
      );
      return null;
    }
  }

  public removeFromLocalStorage(key: string, options?: IStorageOptions): void {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      localStorage.removeItem(fullKey);
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to remove from localStorage: ${key}`,
        error,
      );
    }
  }

  public saveToSessionStorage<T>(
    key: string,
    data: T,
    options?: IStorageOptions,
  ): void {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      const storageData = {
        data,
        timestamp: Date.now(),
      };

      const serialized = JsonUtils.safeStringify(storageData);
      if (serialized) {
        sessionStorage.setItem(fullKey, serialized);
      }
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to save to sessionStorage: ${key}`,
        error,
      );
    }
  }

  public getFromSessionStorage<T>(
    key: string,
    options?: IStorageOptions,
  ): T | null {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      const data = sessionStorage.getItem(fullKey);

      if (!data) return null;

      const parsedData = JsonUtils.safeParse(data);
      return parsedData ? (parsedData.data as T) : null;
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to read from sessionStorage: ${key}`,
        error,
      );
      return null;
    }
  }

  public removeFromSessionStorage(
    key: string,
    options?: IStorageOptions,
  ): void {
    try {
      const fullKey = this.getFullKey(key, options?.userSpecific);
      sessionStorage.removeItem(fullKey);
    } catch (error) {
      logger.warn(
        "StorageService",
        `Failed to remove from sessionStorage: ${key}`,
        error,
      );
    }
  }

  public saveAlerts(alerts: IAlertItem[]): void {
    this.saveToLocalStorage<IAlertItem[]>("AllAlerts", alerts, {
      expirationTime: this.defaultExpirationTime,
    });
  }

  public getAlerts(): IAlertItem[] | null {
    return this.getFromLocalStorage<IAlertItem[]>("AllAlerts");
  }

  public saveDismissedAlerts(alertIds: string[]): void {
    this.saveToSessionStorage<string[]>("DismissedAlerts", alertIds, {
      userSpecific: true,
    });
  }

  public getDismissedAlerts(): string[] {
    return (
      this.getFromSessionStorage<string[]>("DismissedAlerts", {
        userSpecific: true,
      }) || []
    );
  }

  public saveHiddenAlerts(alertIds: string[]): void {
    this.saveToLocalStorage<string[]>("HiddenAlerts", alertIds, {
      userSpecific: true,
    });
  }

  public getHiddenAlerts(): string[] {
    return (
      this.getFromLocalStorage<string[]>("HiddenAlerts", {
        userSpecific: true,
      }) || []
    );
  }

  public clearAllAlertData(): void {
    this.removeFromLocalStorage("AllAlerts");
    this.removeFromSessionStorage("DismissedAlerts", { userSpecific: true });
    this.removeFromLocalStorage("HiddenAlerts", { userSpecific: true });
  }

  public initCrossTabSync(onAlertStorageChange: () => void): () => void {
    const handler = (event: StorageEvent): void => {
      if (!event.key) return;

      const relevantKeys = ["DismissedAlerts", "HiddenAlerts"];
      const isRelevantChange = relevantKeys.some((key) =>
        event.key?.includes(key),
      );

      if (isRelevantChange) {
        logger.debug("StorageService", "Cross-tab storage change detected", {
          key: event.key,
        });
        onAlertStorageChange();
      }
    };

    window.addEventListener("storage", handler);

    return () => {
      window.removeEventListener("storage", handler);
    };
  }

  private getFullKey(key: string, userSpecific?: boolean): string {
    const prefix = "AlertsBanner_";
    const userPrefix = userSpecific && this.userId ? `${this.userId}_` : "";
    return `${prefix}${userPrefix}${key}`;
  }

  private isDataExpired(storageData: Record<string, unknown>): boolean {
    const timestamp = storageData.timestamp as number;
    const expiration = storageData.expiration as number;
    if (!timestamp || !expiration) return false;

    const now = Date.now();
    const expirationTime = timestamp + expiration;

    return now > expirationTime;
  }
}

export default StorageService;
