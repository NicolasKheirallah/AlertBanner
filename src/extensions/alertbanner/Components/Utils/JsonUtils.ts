import { logger } from '../Services/LoggerService';
import { SANITIZATION_CONFIG, VALIDATION_LIMITS } from './AppConstants';

export class JsonUtils {
  public static safeParse<T = any>(jsonString: string | null | undefined): T | null {
    if (!jsonString || typeof jsonString !== 'string') {
      return null;
    }

    try {
      return JSON.parse(jsonString) as T;
    } catch (error) {
      logger.debug('JsonUtils', 'JSON parse failed', { error, jsonPreview: jsonString.substring(0, 50) });
      return null;
    }
  }

  public static safeParseWithDefault<T>(jsonString: string | null | undefined, defaultValue: T): T {
    const result = this.safeParse<T>(jsonString);
    return result !== null ? result : defaultValue;
  }

  public static safeStringify(obj: any, pretty: boolean = false): string | null {
    if (obj === undefined || obj === null) {
      return null;
    }

    try {
      return JSON.stringify(obj, null, pretty ? 2 : 0);
    } catch (error) {
      logger.warn('JsonUtils', 'JSON stringify failed', { error });
      return null;
    }
  }

  public static parseWithValidation<T = any>(
    jsonString: string | null | undefined,
    options: {
      maxDepth?: number;
      checkDangerousKeys?: boolean;
    } = {}
  ): { success: boolean; data: T | null; errors: string[] } {
    const maxDepth = options.maxDepth || VALIDATION_LIMITS.JSON_MAX_DEPTH;
    const checkDangerousKeys = options.checkDangerousKeys !== false;
    const errors: string[] = [];

    if (!jsonString || typeof jsonString !== 'string') {
      errors.push('JSON data is required and must be a string');
      return { success: false, data: null, errors };
    }

    let parsed: any;
    try {
      parsed = JSON.parse(jsonString);
    } catch (parseError) {
      errors.push('Invalid JSON format');
      return { success: false, data: null, errors };
    }

    const depth = this.getObjectDepth(parsed);
    if (depth > maxDepth) {
      errors.push(`JSON structure is too deeply nested (max depth: ${maxDepth}, actual: ${depth})`);
    }

    if (checkDangerousKeys && this.containsDangerousKeys(parsed)) {
      errors.push('JSON contains potentially dangerous property names');
    }

    const success = errors.length === 0;
    return {
      success,
      data: success ? parsed as T : null,
      errors
    };
  }

  public static getObjectDepth(obj: any, currentDepth: number = 0): number {
    if (obj === null || typeof obj !== 'object') {
      return currentDepth;
    }

    if (Array.isArray(obj)) {
      if (obj.length === 0) {
        return currentDepth + 1;
      }
      return Math.max(...obj.map(item => this.getObjectDepth(item, currentDepth + 1)));
    }

    const keys = Object.keys(obj);
    if (keys.length === 0) {
      return currentDepth + 1;
    }

    return Math.max(...keys.map(key => this.getObjectDepth(obj[key], currentDepth + 1)));
  }

  public static containsDangerousKeys(obj: any): boolean {
    if (obj === null || typeof obj !== 'object') {
      return false;
    }

    const dangerousKeys = SANITIZATION_CONFIG.DANGEROUS_JSON_KEYS;

    for (const key of Object.keys(obj)) {
      if (dangerousKeys.includes(key as any)) {
        return true;
      }
    }

    for (const value of Object.values(obj)) {
      if (typeof value === 'object' && value !== null) {
        if (this.containsDangerousKeys(value)) {
          return true;
        }
      }
    }

    return false;
  }

  // Deep clone using JSON parse/stringify - note: won't clone functions, undefined, or circular refs
  public static deepClone<T>(obj: T): T | null {
    const stringified = this.safeStringify(obj);
    if (!stringified) {
      return null;
    }
    return this.safeParse<T>(stringified);
  }

  public static deepEquals(obj1: any, obj2: any): boolean {
    const str1 = this.safeStringify(obj1);
    const str2 = this.safeStringify(obj2);

    if (str1 === null || str2 === null) {
      return obj1 === obj2;
    }

    return str1 === str2;
  }

  public static merge<T extends object>(...objects: Array<Partial<T> | null | undefined>): Partial<T> {
    return objects.reduce<Partial<T>>((acc, obj) => {
      if (obj) {
        return { ...acc, ...obj };
      }
      return acc;
    }, {} as Partial<T>);
  }

  public static pick<T extends object, K extends keyof T>(obj: T, keys: K[]): Pick<T, K> {
    const result = {} as Pick<T, K>;
    keys.forEach(key => {
      if (key in obj) {
        result[key] = obj[key];
      }
    });
    return result;
  }

  public static omit<T extends object, K extends keyof T>(obj: T, keys: K[]): Omit<T, K> {
    const result = { ...obj };
    keys.forEach(key => {
      delete result[key];
    });
    return result;
  }

  public static getNestedValue<T = any>(obj: any, path: string, defaultValue?: T): T | undefined {
    const keys = path.split('.');
    let current = obj;

    for (const key of keys) {
      if (current === null || current === undefined || typeof current !== 'object') {
        return defaultValue;
      }
      current = current[key];
    }

    return current !== undefined ? current : defaultValue;
  }

  public static setNestedValue(obj: any, path: string, value: any): void {
    const keys = path.split('.');
    const lastKey = keys.pop();

    if (!lastKey) {
      return;
    }

    let current = obj;
    for (const key of keys) {
      if (!(key in current) || typeof current[key] !== 'object') {
        current[key] = {};
      }
      current = current[key];
    }

    current[lastKey] = value;
  }

  public static isPlainObject(value: any): boolean {
    return value !== null &&
           typeof value === 'object' &&
           !Array.isArray(value) &&
           !(value instanceof Date) &&
           !(value instanceof RegExp);
  }

  public static flatten(obj: any, prefix: string = ''): Record<string, any> {
    const result: Record<string, any> = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        const newKey = prefix ? `${prefix}.${key}` : key;

        if (this.isPlainObject(obj[key])) {
          Object.assign(result, this.flatten(obj[key], newKey));
        } else {
          result[newKey] = obj[key];
        }
      }
    }

    return result;
  }

  public static unflatten(obj: Record<string, any>): any {
    const result: any = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        this.setNestedValue(result, key, obj[key]);
      }
    }

    return result;
  }

  public static filterObject<T extends object>(
    obj: T,
    predicate: (key: string, value: any) => boolean
  ): Partial<T> {
    const result: any = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key) && predicate(key, obj[key])) {
        result[key] = obj[key];
      }
    }

    return result;
  }

  public static mapObject<T extends object, R>(
    obj: T,
    mapper: (key: string, value: any) => R
  ): Record<string, R> {
    const result: Record<string, R> = {};

    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        result[key] = mapper(key, obj[key]);
      }
    }

    return result;
  }

  public static compact<T extends object>(obj: T): Partial<T> {
    return this.filterObject(obj, (_, value) => value !== null && value !== undefined);
  }

  public static prettyPrint(obj: any, indent: number = 2): string {
    return this.safeStringify(obj, true) || '';
  }
}
