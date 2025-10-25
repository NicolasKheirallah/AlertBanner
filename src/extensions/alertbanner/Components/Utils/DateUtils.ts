/**
 * Date utility functions for AlertBanner
 * Consolidates date formatting, parsing, and calculation logic
 */
export class DateUtils {
  // Time duration constants in milliseconds
  public static readonly MILLISECONDS = {
    SECOND: 1000,
    MINUTE: 60 * 1000,
    HOUR: 60 * 60 * 1000,
    DAY: 24 * 60 * 60 * 1000,
    WEEK: 7 * 24 * 60 * 60 * 1000,
    MONTH: 30 * 24 * 60 * 60 * 1000,
    YEAR: 365 * 24 * 60 * 60 * 1000
  };

  /**
   * Get current timestamp as ISO string
   */
  public static nowISO(): string {
    return new Date().toISOString();
  }

  /**
   * Get current timestamp in milliseconds
   */
  public static nowMillis(): number {
    return Date.now();
  }

  /**
   * Parse date string or Date object to Date
   * @returns Date object or null if invalid
   */
  public static parseDate(date: Date | string | undefined | null): Date | null {
    if (!date) {
      return null;
    }

    const dateObj = typeof date === 'string' ? new Date(date) : date;

    // Check if date is valid
    if (isNaN(dateObj.getTime())) {
      return null;
    }

    return dateObj;
  }

  /**
   * Check if a date string or Date object is valid
   */
  public static isValidDate(date: Date | string | undefined | null): boolean {
    return this.parseDate(date) !== null;
  }

  /**
   * Convert date to ISO string, handling undefined/null
   */
  public static toISOString(date: Date | string | undefined | null): string | undefined {
    const parsed = this.parseDate(date);
    return parsed ? parsed.toISOString() : undefined;
  }

  /**
   * Convert date for HTML datetime-local input (adjusts for timezone offset)
   * HTML datetime-local inputs expect local time in format: YYYY-MM-DDTHH:mm
   */
  public static toDateTimeLocalValue(date: Date | string | undefined | null): string {
    const parsed = this.parseDate(date);
    if (!parsed) {
      return '';
    }

    // Adjust for timezone offset to get local time
    const localDate = new Date(parsed.getTime() - parsed.getTimezoneOffset() * 60000);
    return localDate.toISOString().slice(0, 16);
  }

  /**
   * Add duration to a date
   * @param date - Base date
   * @param amount - Amount to add
   * @param unit - Time unit (seconds, minutes, hours, days, weeks, months, years)
   */
  public static addDuration(
    date: Date | string,
    amount: number,
    unit: 'seconds' | 'minutes' | 'hours' | 'days' | 'weeks' | 'months' | 'years'
  ): Date {
    const baseDate = this.parseDate(date);
    if (!baseDate) {
      throw new Error('Invalid date provided to addDuration');
    }

    const millisToAdd = amount * this.getMillisecondsForUnit(unit);
    return new Date(baseDate.getTime() + millisToAdd);
  }

  /**
   * Add duration and return ISO string
   */
  public static addDurationISO(
    date: Date | string,
    amount: number,
    unit: 'seconds' | 'minutes' | 'hours' | 'days' | 'weeks' | 'months' | 'years'
  ): string {
    return this.addDuration(date, amount, unit).toISOString();
  }

  /**
   * Get milliseconds for a time unit
   */
  private static getMillisecondsForUnit(unit: 'seconds' | 'minutes' | 'hours' | 'days' | 'weeks' | 'months' | 'years'): number {
    switch (unit) {
      case 'seconds': return this.MILLISECONDS.SECOND;
      case 'minutes': return this.MILLISECONDS.MINUTE;
      case 'hours': return this.MILLISECONDS.HOUR;
      case 'days': return this.MILLISECONDS.DAY;
      case 'weeks': return this.MILLISECONDS.WEEK;
      case 'months': return this.MILLISECONDS.MONTH;
      case 'years': return this.MILLISECONDS.YEAR;
    }
  }

  /**
   * Calculate difference between two dates in specified unit
   */
  public static diff(
    date1: Date | string,
    date2: Date | string,
    unit: 'seconds' | 'minutes' | 'hours' | 'days' | 'weeks' | 'months' | 'years' = 'days'
  ): number {
    const d1 = this.parseDate(date1);
    const d2 = this.parseDate(date2);

    if (!d1 || !d2) {
      throw new Error('Invalid dates provided to diff');
    }

    const diffInMillis = d1.getTime() - d2.getTime();
    return diffInMillis / this.getMillisecondsForUnit(unit);
  }

  /**
   * Check if date1 is before date2
   */
  public static isBefore(date1: Date | string, date2: Date | string): boolean {
    const d1 = this.parseDate(date1);
    const d2 = this.parseDate(date2);

    if (!d1 || !d2) {
      return false;
    }

    return d1.getTime() < d2.getTime();
  }

  /**
   * Check if date1 is after date2
   */
  public static isAfter(date1: Date | string, date2: Date | string): boolean {
    const d1 = this.parseDate(date1);
    const d2 = this.parseDate(date2);

    if (!d1 || !d2) {
      return false;
    }

    return d1.getTime() > d2.getTime();
  }

  /**
   * Check if dates are on the same day
   */
  public static isSameDay(date1: Date | string, date2: Date | string): boolean {
    const d1 = this.parseDate(date1);
    const d2 = this.parseDate(date2);

    if (!d1 || !d2) {
      return false;
    }

    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
  }

  /**
   * Check if date is today
   */
  public static isToday(date: Date | string): boolean {
    return this.isSameDay(date, new Date());
  }

  /**
   * Check if date is in the past
   */
  public static isPast(date: Date | string, referenceDate: Date = new Date()): boolean {
    return this.isBefore(date, referenceDate);
  }

  /**
   * Check if date is in the future
   */
  public static isFuture(date: Date | string, referenceDate: Date = new Date()): boolean {
    return this.isAfter(date, referenceDate);
  }

  /**
   * Get start of day (00:00:00.000)
   */
  public static startOfDay(date: Date | string = new Date()): Date {
    const d = this.parseDate(date);
    if (!d) {
      return new Date();
    }

    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  /**
   * Get end of day (23:59:59.999)
   */
  public static endOfDay(date: Date | string = new Date()): Date {
    const d = this.parseDate(date);
    if (!d) {
      return new Date();
    }

    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
  }

  /**
   * Check if date is within a range (inclusive)
   */
  public static isWithinRange(
    date: Date | string,
    startDate: Date | string,
    endDate: Date | string
  ): boolean {
    const d = this.parseDate(date);
    const start = this.parseDate(startDate);
    const end = this.parseDate(endDate);

    if (!d || !start || !end) {
      return false;
    }

    return d.getTime() >= start.getTime() && d.getTime() <= end.getTime();
  }

  /**
   * Generate a unique ID with timestamp prefix
   * Format: {timestamp}-{randomString}
   */
  public static generateTimestampId(length: number = 11): string {
    const randomPart = Math.random().toString(36).substring(2, 2 + length);
    return `${Date.now()}-${randomPart}`;
  }

  /**
   * Format cache age for display
   * @param timestamp - Cache timestamp in milliseconds
   * @returns Human-readable string like "5 minutes ago"
   */
  public static formatCacheAge(timestamp: number): string {
    const ageInMillis = Date.now() - timestamp;

    if (ageInMillis < this.MILLISECONDS.MINUTE) {
      return 'just now';
    } else if (ageInMillis < this.MILLISECONDS.HOUR) {
      const minutes = Math.floor(ageInMillis / this.MILLISECONDS.MINUTE);
      return `${minutes} minute${minutes !== 1 ? 's' : ''} ago`;
    } else if (ageInMillis < this.MILLISECONDS.DAY) {
      const hours = Math.floor(ageInMillis / this.MILLISECONDS.HOUR);
      return `${hours} hour${hours !== 1 ? 's' : ''} ago`;
    } else {
      const days = Math.floor(ageInMillis / this.MILLISECONDS.DAY);
      return `${days} day${days !== 1 ? 's' : ''} ago`;
    }
  }

  /**
   * Check if cached data is still fresh
   * @param timestamp - Cache timestamp in milliseconds
   * @param maxAge - Maximum age in milliseconds
   */
  public static isCacheFresh(timestamp: number, maxAge: number): boolean {
    return (Date.now() - timestamp) < maxAge;
  }

  /**
   * Sort dates ascending (oldest first)
   */
  public static sortAscending(dates: Array<Date | string>): Array<Date | string> {
    return [...dates].sort((a, b) => {
      const d1 = this.parseDate(a);
      const d2 = this.parseDate(b);
      if (!d1 || !d2) return 0;
      return d1.getTime() - d2.getTime();
    });
  }

  /**
   * Sort dates descending (newest first)
   */
  public static sortDescending(dates: Array<Date | string>): Array<Date | string> {
    return [...dates].sort((a, b) => {
      const d1 = this.parseDate(a);
      const d2 = this.parseDate(b);
      if (!d1 || !d2) return 0;
      return d2.getTime() - d1.getTime();
    });
  }

  /**
   * Get relative time description
   * @param date - Date to describe
   * @returns String like "in 2 days" or "3 hours ago"
   */
  public static getRelativeTime(date: Date | string): string {
    const d = this.parseDate(date);
    if (!d) {
      return 'unknown';
    }

    const now = new Date();
    const diffInSeconds = Math.abs(Math.floor((d.getTime() - now.getTime()) / 1000));
    const isPast = d.getTime() < now.getTime();
    const prefix = isPast ? '' : 'in ';
    const suffix = isPast ? ' ago' : '';

    if (diffInSeconds < 60) {
      return isPast ? 'just now' : 'in a moment';
    } else if (diffInSeconds < 3600) {
      const minutes = Math.floor(diffInSeconds / 60);
      return `${prefix}${minutes} minute${minutes !== 1 ? 's' : ''}${suffix}`;
    } else if (diffInSeconds < 86400) {
      const hours = Math.floor(diffInSeconds / 3600);
      return `${prefix}${hours} hour${hours !== 1 ? 's' : ''}${suffix}`;
    } else if (diffInSeconds < 2592000) {
      const days = Math.floor(diffInSeconds / 86400);
      return `${prefix}${days} day${days !== 1 ? 's' : ''}${suffix}`;
    } else if (diffInSeconds < 31536000) {
      const months = Math.floor(diffInSeconds / 2592000);
      return `${prefix}${months} month${months !== 1 ? 's' : ''}${suffix}`;
    } else {
      const years = Math.floor(diffInSeconds / 31536000);
      return `${prefix}${years} year${years !== 1 ? 's' : ''}${suffix}`;
    }
  }
}
