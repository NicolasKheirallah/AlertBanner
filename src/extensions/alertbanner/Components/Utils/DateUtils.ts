export class DateUtils {
  public static readonly MILLISECONDS = {
    SECOND: 1000,
    MINUTE: 60 * 1000,
    HOUR: 60 * 60 * 1000,
    DAY: 24 * 60 * 60 * 1000,
    WEEK: 7 * 24 * 60 * 60 * 1000,
    MONTH: 30 * 24 * 60 * 60 * 1000,
    YEAR: 365 * 24 * 60 * 60 * 1000
  };

  public static nowISO(): string {
    return new Date().toISOString();
  }

  public static nowMillis(): number {
    return Date.now();
  }

  // Parse date string or Date object to Date, returns null if invalid
  public static parseDate(date: Date | string | undefined | null): Date | null {
    if (!date) {
      return null;
    }

    const dateObj = typeof date === 'string' ? new Date(date) : date;

    if (isNaN(dateObj.getTime())) {
      return null;
    }

    return dateObj;
  }

  public static isValidDate(date: Date | string | undefined | null): boolean {
    return this.parseDate(date) !== null;
  }

  public static toISOString(date: Date | string | undefined | null): string | undefined {
    const parsed = this.parseDate(date);
    return parsed ? parsed.toISOString() : undefined;
  }

  public static toDateTimeLocalValue(date: Date | string | undefined | null): string {
    const parsed = this.parseDate(date);
    if (!parsed) {
      return '';
    }

    const localDate = new Date(parsed.getTime() - parsed.getTimezoneOffset() * 60000);
    return localDate.toISOString().slice(0, 16);
  }

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

  // Add duration to a date and return ISO string (used by ListProvisioningService)
  public static addDurationISO(
    date: Date | string,
    amount: number,
    unit: 'seconds' | 'minutes' | 'hours' | 'days' | 'weeks' | 'months' | 'years'
  ): string {
    return this.addDuration(date, amount, unit).toISOString();
  }

  public static formatForDisplay(date: Date | string | undefined | null): string {
    const parsed = this.parseDate(date);
    if (!parsed) {
      return '';
    }

    return parsed.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }

  public static formatDateTimeForDisplay(date: Date | string | undefined | null): string {
    const parsed = this.parseDate(date);
    if (!parsed) {
      return '';
    }

    return parsed.toLocaleString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }
}
