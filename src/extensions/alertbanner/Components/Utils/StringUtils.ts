/**
 * String utility functions for AlertBanner
 * Consolidates common string operations and validations
 */
export class StringUtils {
  /**
   * Trim string or return default value if empty/null/undefined
   */
  public static trimOrDefault(value: string | null | undefined, defaultValue: string = ''): string {
    if (!value || typeof value !== 'string') {
      return defaultValue;
    }
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : defaultValue;
  }

  /**
   * Normalize string for comparison (trim + lowercase)
   */
  public static normalizeForComparison(value: string | null | undefined): string {
    if (!value || typeof value !== 'string') {
      return '';
    }
    return value.trim().toLowerCase();
  }

  /**
   * Case-insensitive string equality check
   */
  public static equalsIgnoreCase(str1: string | null | undefined, str2: string | null | undefined): boolean {
    return this.normalizeForComparison(str1) === this.normalizeForComparison(str2);
  }

  /**
   * Check if string is null, undefined, or empty (after trim)
   */
  public static isEmpty(value: string | null | undefined): boolean {
    return !value || typeof value !== 'string' || value.trim().length === 0;
  }

  /**
   * Check if string is not empty
   */
  public static isNotEmpty(value: string | null | undefined): boolean {
    return !this.isEmpty(value);
  }

  /**
   * Truncate string to max length with ellipsis
   */
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

  /**
   * Capitalize first letter of string
   */
  public static capitalize(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    const str = value!.trim();
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
  }

  /**
   * Capitalize first letter of each word
   */
  public static capitalizeWords(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.trim().split(/\s+/).map(word => this.capitalize(word)).join(' ');
  }

  /**
   * Convert string to kebab-case
   */
  public static toKebabCase(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!
      .trim()
      .replace(/([a-z])([A-Z])/g, '$1-$2') // camelCase to kebab-case
      .replace(/[\s_]+/g, '-') // spaces/underscores to hyphens
      .toLowerCase();
  }

  /**
   * Convert string to camelCase
   */
  public static toCamelCase(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!
      .trim()
      .replace(/[-_\s](.)/g, (_, char) => char.toUpperCase())
      .replace(/^(.)/, (_, char) => char.toLowerCase());
  }

  /**
   * Pluralize word based on count
   */
  public static pluralize(word: string, count: number, plural?: string): string {
    if (count === 1) {
      return word;
    }
    return plural || `${word}s`;
  }

  /**
   * Format count with pluralized word
   * Example: formatCount(1, 'item') => "1 item", formatCount(5, 'item') => "5 items"
   */
  public static formatCount(count: number, word: string, plural?: string): string {
    return `${count} ${this.pluralize(word, count, plural)}`;
  }

  /**
   * Remove special characters (keep only alphanumeric, spaces, hyphens)
   */
  public static sanitizeForId(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!
      .trim()
      .replace(/[^a-zA-Z0-9\s-]/g, '')
      .replace(/\s+/g, '-')
      .toLowerCase();
  }

  /**
   * Check if string starts with prefix (case-insensitive option)
   */
  public static startsWith(
    value: string | null | undefined,
    prefix: string,
    ignoreCase: boolean = false
  ): boolean {
    if (this.isEmpty(value) || !prefix) {
      return false;
    }

    if (ignoreCase) {
      return this.normalizeForComparison(value).startsWith(this.normalizeForComparison(prefix));
    }

    return value!.startsWith(prefix);
  }

  /**
   * Check if string ends with suffix (case-insensitive option)
   */
  public static endsWith(
    value: string | null | undefined,
    suffix: string,
    ignoreCase: boolean = false
  ): boolean {
    if (this.isEmpty(value) || !suffix) {
      return false;
    }

    if (ignoreCase) {
      return this.normalizeForComparison(value).endsWith(this.normalizeForComparison(suffix));
    }

    return value!.endsWith(suffix);
  }

  /**
   * Check if string contains substring (case-insensitive option)
   */
  public static contains(
    value: string | null | undefined,
    substring: string,
    ignoreCase: boolean = false
  ): boolean {
    if (this.isEmpty(value) || !substring) {
      return false;
    }

    if (ignoreCase) {
      return this.normalizeForComparison(value).includes(this.normalizeForComparison(substring));
    }

    return value!.includes(substring);
  }

  /**
   * Split string by delimiter and trim each part
   */
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

  /**
   * Join array of strings with delimiter, filtering out empty values
   */
  public static joinNonEmpty(values: Array<string | null | undefined>, delimiter: string = ', '): string {
    return values
      .filter(v => this.isNotEmpty(v))
      .map(v => v!.trim())
      .join(delimiter);
  }

  /**
   * Escape HTML special characters
   */
  public static escapeHtml(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    const map: { [key: string]: string } = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
      '/': '&#x2F;'
    };

    return value!.replace(/[&<>"'/]/g, char => map[char]);
  }

  /**
   * Strip HTML tags from string
   */
  public static stripHtml(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/<[^>]*>/g, '');
  }

  /**
   * Convert line breaks to <br> tags
   */
  public static nl2br(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/\n/g, '<br>');
  }

  /**
   * Remove extra whitespace (multiple spaces/tabs/newlines to single space)
   */
  public static normalizeWhitespace(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/\s+/g, ' ').trim();
  }

  /**
   * Mask sensitive information (show first/last N chars, mask middle)
   */
  public static mask(
    value: string | null | undefined,
    visibleStart: number = 3,
    visibleEnd: number = 3,
    maskChar: string = '*'
  ): string {
    if (this.isEmpty(value)) {
      return '';
    }

    const str = value!;
    if (str.length <= visibleStart + visibleEnd) {
      return maskChar.repeat(str.length);
    }

    const start = str.substring(0, visibleStart);
    const end = str.substring(str.length - visibleEnd);
    const middle = maskChar.repeat(str.length - visibleStart - visibleEnd);

    return start + middle + end;
  }

  /**
   * Generate random alphanumeric string
   */
  public static randomAlphanumeric(length: number): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }

  /**
   * Check if string is valid email format (basic check)
   */
  public static isEmail(value: string | null | undefined): boolean {
    if (this.isEmpty(value)) {
      return false;
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(value!.trim());
  }

  /**
   * Check if string is valid URL format
   */
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

  /**
   * Extract domain from URL
   */
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

  /**
   * Format file size in bytes to human-readable format
   */
  public static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return `${Math.round(bytes / Math.pow(k, i) * 100) / 100} ${sizes[i]}`;
  }

  /**
   * Highlight search term in text (wrap with <mark> tags)
   */
  public static highlightSearchTerm(
    text: string | null | undefined,
    searchTerm: string | null | undefined,
    ignoreCase: boolean = true
  ): string {
    if (this.isEmpty(text) || this.isEmpty(searchTerm)) {
      return text || '';
    }

    const flags = ignoreCase ? 'gi' : 'g';
    const regex = new RegExp(`(${searchTerm!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, flags);

    return text!.replace(regex, '<mark>$1</mark>');
  }
}
