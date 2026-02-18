import { logger } from '../Services/LoggerService';
import { VALIDATION_LIMITS } from './AppConstants';

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

    const success = errors.length === 0;
    return {
      success,
      data: success ? parsed as T : null,
      errors
    };
  }

  private static getObjectDepth(obj: any, currentDepth: number = 0): number {
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

  // Deep clone using JSON parse/stringify - note: won't clone functions, undefined, or circular refs
  public static deepClone<T>(obj: T): T | null {
    const stringified = this.safeStringify(obj);
    if (!stringified) {
      return null;
    }
    return this.safeParse<T>(stringified);
  }
}
