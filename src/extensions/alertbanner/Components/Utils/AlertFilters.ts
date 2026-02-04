import { SiteIdUtils } from './SiteIdUtils';
import { IAlertItem, ContentType, AlertPriority, NotificationType, TargetLanguage } from '../Alerts/IAlerts';
import { logger } from '../Services/LoggerService';

/**
 * Utility class for filtering alerts
 * Consolidates duplicate filtering logic from multiple components
 */
export class AlertFilters {
  /**
   * Remove duplicate alerts based on ID
   */
  public static removeDuplicates(alerts: IAlertItem[]): IAlertItem[] {
    const seenIds = new Set<string>();
    const duplicates: string[] = [];

    const unique = alerts.filter((alert) => {
      if (seenIds.has(alert.id)) {
        duplicates.push(alert.id);
        return false;
      } else {
        seenIds.add(alert.id);
        return true;
      }
    });

    if (duplicates.length > 0) {
      logger.debug('AlertFilters', `Removed ${duplicates.length} duplicate alerts`, { duplicateIds: duplicates });
    }

    return unique;
  }

  /**
   * Filter out dismissed alerts
   */
  public static filterDismissed(alerts: IAlertItem[], dismissedIds: string[]): IAlertItem[] {
    if (dismissedIds.length === 0) {
      return alerts;
    }

    const dismissedSet = new Set(dismissedIds);
    return alerts.filter(alert => !dismissedSet.has(alert.id));
  }

  /**
   * Filter out hidden alerts
   */
  public static filterHidden(alerts: IAlertItem[], hiddenIds: string[]): IAlertItem[] {
    if (hiddenIds.length === 0) {
      return alerts;
    }

    const hiddenSet = new Set(hiddenIds);
    return alerts.filter(alert => !hiddenSet.has(alert.id));
  }

  /**
   * Filter alerts by dismissed and hidden IDs
   */
  public static filterDismissedAndHidden(
    alerts: IAlertItem[],
    dismissedIds: string[],
    hiddenIds: string[]
  ): IAlertItem[] {
    if (dismissedIds.length === 0 && hiddenIds.length === 0) {
      return alerts;
    }

    const dismissedSet = new Set(dismissedIds);
    const hiddenSet = new Set(hiddenIds);

    return alerts.filter(alert =>
      !dismissedSet.has(alert.id) && !hiddenSet.has(alert.id)
    );
  }

  /**
   * Filter alerts by target sites
   * If alert has no targetSites, it's available to all sites
   */
  public static filterByTargetSites(alerts: IAlertItem[], scopedSiteIds: string[]): IAlertItem[] {
    if (scopedSiteIds.length === 0) {
      return alerts;
    }

    // Build unique set of all scoped site variations
    const scopeSet = new Set<string>();
    scopedSiteIds.forEach(site => {
      SiteIdUtils.generateSiteVariations(site).forEach(v => scopeSet.add(v));
    });

    return alerts.filter(alert => {
      // If no target sites specified, show to all sites
      if (!alert.targetSites || alert.targetSites.length === 0) {
        return true;
      }

      // Check if any of the alert's target sites match the scoped sites
      return alert.targetSites.some(targetSiteId => {
        const variations = SiteIdUtils.generateSiteVariations(String(targetSiteId));
        return variations.some(v => scopeSet.has(v));
      });
    });
  }

  /**
   * Check if alert is a template
   */
  public static isTemplate(alert: IAlertItem): boolean {
    return alert.contentType === ContentType.Template ||
           alert.AlertType?.toLowerCase().includes('template') ||
           alert.title?.toLowerCase().includes('template');
  }

  /**
   * Check if alert is a draft
   */
  public static isDraft(alert: IAlertItem): boolean {
    return alert.contentType === ContentType.Draft ||
           alert.status?.toLowerCase() === 'draft';
  }

  /**
   * Check if alert is auto-saved
   */
  public static isAutoSaved(alert: IAlertItem): boolean {
    return alert.title?.startsWith('[Auto-saved]') ||
           alert.title?.startsWith('[auto-saved]');
  }

  /**
   * Filter out templates
   */
  public static excludeTemplates(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isTemplate(alert));
  }

  /**
   * Filter out drafts
   */
  public static excludeDrafts(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isDraft(alert));
  }

  /**
   * Filter out auto-saved items
   */
  public static excludeAutoSaved(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isAutoSaved(alert));
  }

  /**
   * Filter out templates, drafts, and auto-saved items (common exclusion pattern)
   * Alerts with empty/null ItemType are treated as valid alerts (backward compatibility)
   */
  public static excludeNonPublicAlerts(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => {
      // If contentType is not set, treat as a regular alert (backward compatibility)
      const isTemplate = alert.contentType === ContentType.Template;
      const isDraft = alert.contentType === ContentType.Draft ||
                      alert.status?.toLowerCase() === 'draft';
      const isAutoSaved = this.isAutoSaved(alert);
      const isSettings = alert.contentType === 'settings' as ContentType;
      
      return !isTemplate && !isDraft && !isAutoSaved && !isSettings;
    });
  }

  /**
   * Filter alerts by content type
   */
  public static filterByContentType(alerts: IAlertItem[], contentType: ContentType | 'all'): IAlertItem[] {
    if (contentType === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.contentType === contentType);
  }

  /**
   * Filter alerts by status
   */
  public static filterByStatus(
    alerts: IAlertItem[],
    status: 'Active' | 'Expired' | 'Scheduled' | 'Draft' | 'all'
  ): IAlertItem[] {
    if (status === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.status?.toLowerCase() === status.toLowerCase());
  }

  /**
   * Filter alerts by priority
   */
  public static filterByPriority(alerts: IAlertItem[], priority: AlertPriority | 'all'): IAlertItem[] {
    if (priority === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.priority === priority);
  }

  /**
   * Filter alerts by alert type
   */
  public static filterByAlertType(alerts: IAlertItem[], alertType: string): IAlertItem[] {
    if (alertType === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.AlertType === alertType);
  }

  /**
   * Filter alerts by target language
   */
  public static filterByLanguage(alerts: IAlertItem[], language: TargetLanguage | 'all'): IAlertItem[] {
    if (language === 'all') {
      return alerts;
    }
    // Include alerts that match the language OR have targetLanguage set to 'all'
    return alerts.filter(alert => 
      alert.targetLanguage?.toLowerCase() === language.toLowerCase() ||
      alert.targetLanguage?.toLowerCase() === 'all'
    );
  }

  /**
   * Filter alerts by notification type
   */
  public static filterByNotificationType(
    alerts: IAlertItem[],
    notificationType: NotificationType | 'all'
  ): IAlertItem[] {
    if (notificationType === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.notificationType === notificationType);
  }

  /**
   * Check if alert is currently active based on schedule
   */
  public static isActive(alert: IAlertItem, currentTime: Date = new Date()): boolean {
    // If scheduledStart exists and is in the future, not yet active
    if (alert.scheduledStart && new Date(alert.scheduledStart) > currentTime) {
      return false;
    }

    // If scheduledEnd exists and is in the past, already expired
    if (alert.scheduledEnd && new Date(alert.scheduledEnd) < currentTime) {
      return false;
    }

    // Alert is active if status is 'Active' or 'Scheduled' with passed start time
    return alert.status === 'Active' ||
      (alert.status === 'Scheduled' &&
        (!alert.scheduledStart || new Date(alert.scheduledStart) <= currentTime));
  }

  /**
   * Filter alerts that are currently active
   */
  public static filterActiveAlerts(alerts: IAlertItem[], currentTime: Date = new Date()): IAlertItem[] {
    return alerts.filter(alert => this.isActive(alert, currentTime));
  }

  /**
   * Check if alert should be expired based on schedule
   */
  public static shouldBeExpired(alert: IAlertItem, currentTime: Date = new Date()): boolean {
    return alert.scheduledEnd !== undefined &&
           new Date(alert.scheduledEnd) < currentTime &&
           alert.status !== 'Expired';
  }

  /**
   * Check if alert should transition from Scheduled to Active
   */
  public static shouldBeActivated(alert: IAlertItem, currentTime: Date = new Date()): boolean {
    return alert.scheduledStart !== undefined &&
           new Date(alert.scheduledStart) <= currentTime &&
           alert.status === 'Scheduled';
  }

  /**
   * Filter alerts by date range
   */
  public static filterByDateRange(
    alerts: IAlertItem[],
    fromDate?: Date,
    toDate?: Date,
    dateField: 'createdDate' | 'scheduledStart' | 'scheduledEnd' = 'createdDate'
  ): IAlertItem[] {
    return alerts.filter(alert => {
      const dateValue = alert[dateField];
      if (!dateValue) {
        return false;
      }

      const alertDate = new Date(dateValue);

      if (fromDate && alertDate < fromDate) {
        return false;
      }

      if (toDate && alertDate > toDate) {
        return false;
      }

      return true;
    });
  }

  /**
   * Filter alerts created today
   */
  public static filterCreatedToday(alerts: IAlertItem[]): IAlertItem[] {
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    return this.filterByDateRange(alerts, todayStart, todayEnd, 'createdDate');
  }

  /**
   * Filter alerts created in the last N days
   */
  public static filterCreatedInLastDays(alerts: IAlertItem[], days: number): IAlertItem[] {
    const now = new Date();
    const fromDate = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);

    return this.filterByDateRange(alerts, fromDate, now, 'createdDate');
  }

  /**
   * Full-text search across alert fields
   */
  public static searchAlerts(alerts: IAlertItem[], searchTerm: string): IAlertItem[] {
    if (!searchTerm || searchTerm.trim() === '') {
      return alerts;
    }

    const search = searchTerm.toLowerCase().trim();

    return alerts.filter(alert =>
      alert.title?.toLowerCase().includes(search) ||
      alert.description?.toLowerCase().includes(search) ||
      alert.AlertType?.toLowerCase().includes(search) ||
      alert.createdBy?.toLowerCase().includes(search) ||
      alert.linkDescription?.toLowerCase().includes(search) ||
      alert.contentType?.toLowerCase().includes(search) ||
      alert.priority?.toLowerCase().includes(search) ||
      alert.status?.toLowerCase().includes(search)
    );
  }

  /**
   * Build GraphQL/OData filter string for server-side filtering
   * Useful for SharePoint list queries
   */
  public static buildODataFilter(options: {
    excludeTemplates?: boolean;
    excludeDrafts?: boolean;
    includeScheduledOnly?: boolean;
    dateTime?: string; // ISO format
  }): string {
    const filters: string[] = [];
    const dateTimeNow = options.dateTime || new Date().toISOString();

    // Schedule filtering
    if (options.includeScheduledOnly) {
      filters.push(
        `(fields/ScheduledStart le '${dateTimeNow}' or fields/ScheduledStart eq null)`,
        `(fields/ScheduledEnd ge '${dateTimeNow}' or fields/ScheduledEnd eq null)`
      );
    }

    // Exclude templates
    if (options.excludeTemplates) {
      filters.push(`(fields/ItemType ne 'template')`);
    }

    // Exclude drafts
    if (options.excludeDrafts) {
      filters.push(`(fields/ItemType ne 'draft')`);
    }

    return filters.length > 0 ? filters.join(' and ') : '';
  }

  /**
   * Combine multiple filter functions
   * Useful for chaining filters efficiently
   */
  public static applyMultipleFilters(
    alerts: IAlertItem[],
    filters: Array<(alerts: IAlertItem[]) => IAlertItem[]>
  ): IAlertItem[] {
    return filters.reduce((filtered, filterFn) => filterFn(filtered), alerts);
  }
}
