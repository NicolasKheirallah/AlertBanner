import { logger } from '../Services/LoggerService';

export class ArrayUtils {
  public static unique<T>(array: T[] | null | undefined): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return [...new Set(array)];
  }

  public static uniqueBy<T>(array: T[] | null | undefined, key: keyof T): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    const seen = new Set<any>();
    return array.filter((item) => {
      const value = item[key];
      if (seen.has(value)) {
        return false;
      }
      seen.add(value);
      return true;
    });
  }

  public static uniqueBySelector<T, K>(
    array: T[] | null | undefined,
    selector: (item: T) => K
  ): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    const seen = new Set<K>();
    return array.filter((item) => {
      const value = selector(item);
      if (seen.has(value)) {
        return false;
      }
      seen.add(value);
      return true;
    });
  }

  public static intersection<T>(arr1: T[] | null | undefined, arr2: T[] | null | undefined): T[] {
    if (!arr1 || !arr2 || !Array.isArray(arr1) || !Array.isArray(arr2)) {
      return [];
    }

    const set2 = new Set(arr2);
    return arr1.filter((item) => set2.has(item));
  }

  public static difference<T>(arr1: T[] | null | undefined, arr2: T[] | null | undefined): T[] {
    if (!arr1 || !Array.isArray(arr1)) {
      return [];
    }

    if (!arr2 || !Array.isArray(arr2)) {
      return [...arr1];
    }

    const set2 = new Set(arr2);
    return arr1.filter((item) => !set2.has(item));
  }

  public static groupBy<T>(array: T[] | null | undefined, key: keyof T): Map<any, T[]> {
    const result = new Map<any, T[]>();

    if (!array || !Array.isArray(array)) {
      return result;
    }

    for (const item of array) {
      const groupKey = item[key];
      const group = result.get(groupKey);

      if (group) {
        group.push(item);
      } else {
        result.set(groupKey, [item]);
      }
    }

    return result;
  }

  public static groupBySelector<T, K>(
    array: T[] | null | undefined,
    selector: (item: T) => K
  ): Map<K, T[]> {
    const result = new Map<K, T[]>();

    if (!array || !Array.isArray(array)) {
      return result;
    }

    for (const item of array) {
      const groupKey = selector(item);
      const group = result.get(groupKey);

      if (group) {
        group.push(item);
      } else {
        result.set(groupKey, [item]);
      }
    }

    return result;
  }

  public static chunk<T>(array: T[] | null | undefined, size: number): T[][] {
    if (!array || !Array.isArray(array) || size <= 0) {
      return [];
    }

    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }

    return chunks;
  }

  public static flatten<T>(arrays: T[][] | null | undefined): T[] {
    if (!arrays || !Array.isArray(arrays)) {
      return [];
    }

    return arrays.reduce((acc, arr) => {
      if (Array.isArray(arr)) {
        return acc.concat(arr);
      }
      return acc;
    }, [] as T[]);
  }

  public static flattenDeep(array: any[] | null | undefined): any[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return array.reduce((acc, val) => {
      if (Array.isArray(val)) {
        return acc.concat(this.flattenDeep(val));
      }
      return acc.concat(val);
    }, []);
  }

  public static isEmpty(array: any[] | null | undefined): boolean {
    return !array || !Array.isArray(array) || array.length === 0;
  }

  public static isNotEmpty(array: any[] | null | undefined): boolean {
    return !this.isEmpty(array);
  }

  public static first<T>(array: T[] | null | undefined, defaultValue?: T): T | undefined {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return defaultValue;
    }

    return array[0];
  }

  public static last<T>(array: T[] | null | undefined, defaultValue?: T): T | undefined {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return defaultValue;
    }

    return array[array.length - 1];
  }

  public static at<T>(array: T[] | null | undefined, index: number, defaultValue?: T): T | undefined {
    if (!array || !Array.isArray(array) || index < 0 || index >= array.length) {
      return defaultValue;
    }

    return array[index];
  }

  public static partition<T>(
    array: T[] | null | undefined,
    predicate: (item: T) => boolean
  ): [T[], T[]] {
    const truthy: T[] = [];
    const falsy: T[] = [];

    if (!array || !Array.isArray(array)) {
      return [truthy, falsy];
    }

    for (const item of array) {
      if (predicate(item)) {
        truthy.push(item);
      } else {
        falsy.push(item);
      }
    }

    return [truthy, falsy];
  }

  public static compact<T>(array: (T | null | undefined)[] | null | undefined): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return array.filter((item): item is T => item != null);
  }

  public static sample<T>(array: T[] | null | undefined): T | undefined {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return undefined;
    }

    const randomIndex = Math.floor(Math.random() * array.length);
    return array[randomIndex];
  }

  public static sampleSize<T>(array: T[] | null | undefined, n: number): T[] {
    if (!array || !Array.isArray(array) || n <= 0) {
      return [];
    }

    const shuffled = [...array].sort(() => Math.random() - 0.5);
    return shuffled.slice(0, Math.min(n, array.length));
  }

  public static shuffle<T>(array: T[] | null | undefined): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    const result = [...array];
    for (let i = result.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [result[i], result[j]] = [result[j], result[i]];
    }

    return result;
  }

  public static sortBy<T>(
    array: T[] | null | undefined,
    key: keyof T,
    order: 'asc' | 'desc' = 'asc'
  ): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return [...array].sort((a, b) => {
      const aVal = a[key];
      const bVal = b[key];

      if (aVal < bVal) return order === 'asc' ? -1 : 1;
      if (aVal > bVal) return order === 'asc' ? 1 : -1;
      return 0;
    });
  }

  public static sortBySelector<T, K>(
    array: T[] | null | undefined,
    selector: (item: T) => K,
    order: 'asc' | 'desc' = 'asc'
  ): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return [...array].sort((a, b) => {
      const aVal = selector(a);
      const bVal = selector(b);

      if (aVal < bVal) return order === 'asc' ? -1 : 1;
      if (aVal > bVal) return order === 'asc' ? 1 : -1;
      return 0;
    });
  }

  public static countBy<T>(array: T[] | null | undefined, key: keyof T): Map<any, number> {
    const result = new Map<any, number>();

    if (!array || !Array.isArray(array)) {
      return result;
    }

    for (const item of array) {
      const keyValue = item[key];
      result.set(keyValue, (result.get(keyValue) || 0) + 1);
    }

    return result;
  }

  public static findIndex<T>(
    array: T[] | null | undefined,
    predicate: (item: T) => boolean
  ): number {
    if (!array || !Array.isArray(array)) {
      return -1;
    }

    return array.findIndex(predicate);
  }

  public static findLastIndex<T>(
    array: T[] | null | undefined,
    predicate: (item: T) => boolean
  ): number {
    if (!array || !Array.isArray(array)) {
      return -1;
    }

    for (let i = array.length - 1; i >= 0; i--) {
      if (predicate(array[i])) {
        return i;
      }
    }

    return -1;
  }

  public static contains<T>(array: T[] | null | undefined, item: T): boolean {
    if (!array || !Array.isArray(array)) {
      return false;
    }

    return array.includes(item);
  }

  public static areEqual<T>(arr1: T[] | null | undefined, arr2: T[] | null | undefined): boolean {
    if (arr1 === arr2) {
      return true;
    }

    if (!arr1 || !arr2 || !Array.isArray(arr1) || !Array.isArray(arr2)) {
      return false;
    }

    if (arr1.length !== arr2.length) {
      return false;
    }

    return arr1.every((item, index) => item === arr2[index]);
  }

  public static sum(array: number[] | null | undefined): number {
    if (!array || !Array.isArray(array)) {
      return 0;
    }

    return array.reduce((sum, val) => sum + (typeof val === 'number' ? val : 0), 0);
  }

  public static average(array: number[] | null | undefined): number {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return 0;
    }

    return this.sum(array) / array.length;
  }

  public static min(array: number[] | null | undefined): number | undefined {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return undefined;
    }

    return Math.min(...array);
  }

  public static max(array: number[] | null | undefined): number | undefined {
    if (!array || !Array.isArray(array) || array.length === 0) {
      return undefined;
    }

    return Math.max(...array);
  }

  public static range(start: number, end: number, step: number = 1): number[] {
    if (step === 0) {
      logger.warn('ArrayUtils', 'Step cannot be zero, using 1 instead');
      step = 1;
    }

    const result: number[] = [];
    if (step > 0) {
      for (let i = start; i < end; i += step) {
        result.push(i);
      }
    } else {
      for (let i = start; i > end; i += step) {
        result.push(i);
      }
    }

    return result;
  }

  public static zip<T1, T2>(arr1: T1[], arr2: T2[]): [T1, T2][] {
    if (!arr1 || !arr2 || !Array.isArray(arr1) || !Array.isArray(arr2)) {
      return [];
    }

    const length = Math.min(arr1.length, arr2.length);
    const result: [T1, T2][] = [];

    for (let i = 0; i < length; i++) {
      result.push([arr1[i], arr2[i]]);
    }

    return result;
  }

  public static toArray<T>(value: T | T[] | null | undefined): T[] {
    if (value == null) {
      return [];
    }

    if (Array.isArray(value)) {
      return value;
    }

    return [value];
  }
}
