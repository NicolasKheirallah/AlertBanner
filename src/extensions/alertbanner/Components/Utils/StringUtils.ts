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

  public static toCamelCase(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!
      .trim()
      .replace(/[-_\s](.)/g, (_, char) => char.toUpperCase())
      .replace(/^(.)/, (_, char) => char.toLowerCase());
  }

  public static pluralize(word: string, count: number, plural?: string): string {
    if (count === 1) {
      return word;
    }
    return plural || `${word}s`;
  }

  public static formatCount(count: number, word: string, plural?: string): string {
    return `${count} ${this.pluralize(word, count, plural)}`;
  }

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

  public static joinNonEmpty(values: Array<string | null | undefined>, delimiter: string = ', '): string {
    return values
      .filter(v => this.isNotEmpty(v))
      .map(v => v!.trim())
      .join(delimiter);
  }

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

  public static stripHtml(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/<[^>]*>/g, '');
  }

  public static nl2br(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/\n/g, '<br>');
  }

  public static normalizeWhitespace(value: string | null | undefined): string {
    if (this.isEmpty(value)) {
      return '';
    }

    return value!.replace(/\s+/g, ' ').trim();
  }

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

  public static randomAlphanumeric(length: number): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }

  public static isEmail(value: string | null | undefined): boolean {
    if (this.isEmpty(value)) {
      return false;
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(value!.trim());
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

  public static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return `${Math.round(bytes / Math.pow(k, i) * 100) / 100} ${sizes[i]}`;
  }

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

  public static resolveUrl(url?: string | null): string {
    if (this.isEmpty(url)) {
      return '#';
    }

    const trimmedUrl = url!.trim();

    // Already an absolute URL
    if (this.isUrl(trimmedUrl)) {
      return trimmedUrl;
    }

    // Server-side rendering - return as-is
    if (typeof window === 'undefined') {
      return trimmedUrl;
    }

    // Relative URL - prepend origin
    return `${window.location.origin}${trimmedUrl.startsWith('/') ? '' : '/'}${trimmedUrl}`;
  }
}
