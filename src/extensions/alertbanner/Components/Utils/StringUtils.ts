export class StringUtils {
  public static trimOrDefault(value: string | null | undefined, defaultValue: string = ''): string {
    if (!value || typeof value !== 'string') {
      return defaultValue;
    }
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : defaultValue;
  }

  public static normalizeForComparison(value: string | null | undefined): string {
    if (!value || typeof value !== 'string') {
      return '';
    }
    return value.trim().toLowerCase();
  }

  public static equalsIgnoreCase(str1: string | null | undefined, str2: string | null | undefined): boolean {
    return this.normalizeForComparison(str1) === this.normalizeForComparison(str2);
  }

  public static isEmpty(value: string | null | undefined): boolean {
    return !value || typeof value !== 'string' || value.trim().length === 0;
  }

  public static isNotEmpty(value: string | null | undefined): boolean {
    return !this.isEmpty(value);
  }

  public static truncate(value: string | null | undefined, maxLength: number, ellipsis: string = '...'): string {
    if (this.isEmpty(value)) {
      return '';
    }

    const str = value!.trim();
    if (str.length <= maxLength) {
      return str;
    }

    return str.substring(0, maxLength - ellipsis.length) + ellipsis;
  }

  public static capitalize(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    const str = value!.trim();
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  }

  public static capitalizeWords(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.trim().split(/\s+/).map(word => this.capitalize(word)).join(' ');
  }

  public static splitAndTrim(
    value: string | null | undefined,
    delimiter: string = ','
  ): string[] {
    if (this.isEmpty(value)) {
      return [];
    }

    return value!
      .split(delimiter)
      .map(part => part.trim())
      .filter(part => part.length > 0);
  }

  public static isUrl(value: string | null | undefined): boolean {
    if (this.isEmpty(value)) {
      return false;
    }

    try {
      new URL(value!.trim());
      return true;
    } catch {
      return false;
    }
  }

  public static extractDomain(url: string | null | undefined): string | null {
    if (this.isEmpty(url)) {
      return null;
    }

    try {
      const urlObj = new URL(url!.trim());
      return urlObj.hostname;
    } catch {
      return null;
    }
  }

  // Resolve a server-relative URL to an absolute URL
  public static resolveUrl(serverRelativeUrl?: string): string {
    if (!serverRelativeUrl) {
      return '#';
    }

    if (/^https?:\/\//i.test(serverRelativeUrl)) {
      return serverRelativeUrl;
    }

    if (typeof window === 'undefined') {
      return serverRelativeUrl;
    }

    return `${window.location.origin}${serverRelativeUrl}`;
  }

  // Sanitize a string for use as an ID (used by CreateAlertTab)
  public static sanitizeForId(value: string | null | undefined): string {
    if (!value || typeof value !== 'string') {
      return '';
    }

    return value
      .trim()
      .replace(/[^a-zA-Z0-9\s_-]/g, '')
      .replace(/\s+/g, '_')
      .substring(0, 50);
  }
}
