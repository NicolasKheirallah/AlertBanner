export class ArrayUtils {
  public static unique<T>(array: T[] | null | undefined): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return [...new Set(array)];
  }

  public static isEmpty(array: any[] | null | undefined): boolean {
    return !array || !Array.isArray(array) || array.length === 0;
  }

  public static compact<T>(array: (T | null | undefined)[] | null | undefined): T[] {
    if (!array || !Array.isArray(array)) {
      return [];
    }

    return array.filter((item): item is T => item != null);
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
}
