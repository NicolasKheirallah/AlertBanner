import { SiteIdUtils } from './SiteIdUtils';
import { IAlertItem, ContentType, AlertPriority, TargetLanguage } from '../Alerts/IAlerts';
import { logger } from '../Services/LoggerService';

// Utility class for filtering alerts
export class AlertFilters {
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

  public static filterDismissed(alerts: IAlertItem[], dismissedIds: string[]): IAlertItem[] {
    if (dismissedIds.length === 0) {
      return alerts;
    }

    const dismissedSet = new Set(dismissedIds);
    return alerts.filter(alert => !dismissedSet.has(alert.id));
  }

  public static filterHidden(alerts: IAlertItem[], hiddenIds: string[]): IAlertItem[] {
    if (hiddenIds.length === 0) {
      return alerts;
    }

    const hiddenSet = new Set(hiddenIds);
    return alerts.filter(alert => !hiddenSet.has(alert.id));
  }

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

  // Filter alerts by target sites - if alert has no targetSites, it's available to all sites
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

  public static isTemplate(alert: IAlertItem): boolean {
    return alert.contentType === ContentType.Template ||
           alert.AlertType?.toLowerCase().includes('template') ||
           alert.title?.toLowerCase().includes('template');
  }

  public static isDraft(alert: IAlertItem): boolean {
    return alert.contentType === ContentType.Draft ||
           alert.status?.toLowerCase() === 'draft';
  }

  public static isAutoSaved(alert: IAlertItem): boolean {
    return alert.title?.startsWith('[Auto-saved]') ||
           alert.title?.startsWith('[auto-saved]');
  }

  public static excludeTemplates(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isTemplate(alert));
  }

  public static excludeDrafts(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isDraft(alert));
  }

  public static excludeAutoSaved(alerts: IAlertItem[]): IAlertItem[] {
    return alerts.filter(alert => !this.isAutoSaved(alert));
  }

  // Filter out templates, drafts, and auto-saved items (common exclusion pattern)
  // Alerts with empty/null ItemType are treated as valid alerts (backward compatibility)
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

  public static filterByContentType(alerts: IAlertItem[], contentType: ContentType | 'all'): IAlertItem[] {
    if (contentType === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.contentType === contentType);
  }

  public static filterByStatus(
    alerts: IAlertItem[],
    status: 'Active' | 'Expired' | 'Scheduled' | 'Draft' | 'all'
  ): IAlertItem[] {
    if (status === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.status?.toLowerCase() === status.toLowerCase());
  }

  public static filterByPriority(alerts: IAlertItem[], priority: AlertPriority | 'all'): IAlertItem[] {
    if (priority === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.priority === priority);
  }

  public static filterByAlertType(alerts: IAlertItem[], alertType: string): IAlertItem[] {
    if (alertType === 'all') {
      return alerts;
    }
    return alerts.filter(alert => alert.AlertType === alertType);
  }

  // Filter alerts by target language - includes alerts that match the language OR have targetLanguage set to 'all'
  public static filterByLanguage(alerts: IAlertItem[], language: TargetLanguage | 'all'): IAlertItem[] {
    if (language === 'all') {
      return alerts;
    }
    return alerts.filter(alert => 
      alert.targetLanguage?.toLowerCase() === language.toLowerCase() ||
      alert.targetLanguage?.toLowerCase() === 'all'
    );
  }

  // Check if alert is currently active based on schedule
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

  public static filterActiveAlerts(alerts: IAlertItem[], currentTime: Date = new Date()): IAlertItem[] {
    return alerts.filter(alert => this.isActive(alert, currentTime));
  }

  // Full-text search across alert fields
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

}
