import { IAlertItem, IAlertListItem } from '../Services/SharePointAlertService';
import { AlertPriority, NotificationType, ContentType, TargetLanguage, IPersonField } from '../Alerts/IAlerts';
import { logger } from '../Services/LoggerService';
import { JsonUtils } from './JsonUtils';

/**
 * Utility class for transforming SharePoint items to alert objects
 * Consolidates duplicate transformation logic from AlertsContext and SharePointAlertService
 */
export class AlertTransformers {
  /**
   * Parse target sites field from SharePoint
   * Handles arrays, JSON strings, and CSV strings
   */
  public static parseTargetSitesField(raw: any): string[] {
    if (!raw) {
      return [];
    }

    if (Array.isArray(raw)) {
      return raw.map(String).filter(Boolean);
    }

    if (typeof raw === 'string') {
      const trimmed = raw.trim();
      if (!trimmed) {
        return [];
      }

      // Try parsing as JSON first
      const parsed = JsonUtils.safeParse<string[]>(trimmed);
      if (parsed && Array.isArray(parsed)) {
        return parsed.map((entry) => String(entry)).filter(Boolean);
      }

      // Fall back to CSV parsing
      return trimmed.split(',').map(site => site.trim()).filter(Boolean);
    }

    return [];
  }

  /**
   * Map Person field data from SharePoint to IPersonField
   * Handles multiple SharePoint field formats
   */
  public static mapPersonFieldData(personField: any, isGroup: boolean): IPersonField {
    // Handle SharePoint lookup field format
    if (personField.LookupId && personField.LookupValue) {
      return {
        id: personField.LookupId.toString(),
        displayName: personField.LookupValue,
        isGroup: isGroup
      };
    }

    // Handle SharePoint user field format with ID
    if (personField.ID || personField.id) {
      return {
        id: (personField.ID || personField.id).toString(),
        displayName: personField.Title || personField.displayName || '',
        email: personField.EMail || personField.email,
        loginName: personField.Name || personField.loginName,
        isGroup: isGroup
      };
    }

    // Handle Graph API format
    return {
      id: personField.id?.toString() || "",
      displayName: personField.displayName || personField.title || "",
      email: personField.email || personField.mail || "",
      loginName: personField.loginName || personField.userPrincipalName || "",
      isGroup: isGroup
    };
  }

  /**
   * Map target users from SharePoint field
   * Handles both single user and array of users
   */
  public static mapTargetUsers(targetUsersField: any): IPersonField[] {
    if (!targetUsersField) {
      return [];
    }

    try {
      if (Array.isArray(targetUsersField)) {
        return targetUsersField.map((user: any) =>
          this.mapPersonFieldData(user, user.isGroup || false)
        );
      } else {
        // Handle single user case
        return [this.mapPersonFieldData(targetUsersField, targetUsersField.isGroup || false)];
      }
    } catch (error) {
      logger.warn('AlertTransformers', 'Error processing target users', error);
      return [];
    }
  }

  /**
   * Parse alert priority from SharePoint field
   */
  public static parsePriority(priorityField: any): AlertPriority {
    if (!priorityField) {
      return AlertPriority.Medium;
    }

    try {
      const priority = AlertPriority[priorityField as keyof typeof AlertPriority];
      return priority || AlertPriority.Medium;
    } catch {
      return AlertPriority.Medium;
    }
  }

  /**
   * Determine content type from SharePoint fields
   * Checks both ItemType and ContentType fields
   */
  public static parseContentType(fields: any): ContentType {
    // Check ItemType field first (preferred)
    if (fields.ItemType) {
      const itemType = fields.ItemType.toLowerCase();
      if (itemType === 'template') {
        return ContentType.Template;
      } else if (itemType === 'alert') {
        return ContentType.Alert;
      } else if (itemType === 'draft') {
        return ContentType.Draft;
      }
    }

    // Fall back to ContentType field
    if (fields.ContentType) {
      const contentType = fields.ContentType.toLowerCase();
      if (contentType === 'template') {
        return ContentType.Template;
      } else if (contentType === 'draft') {
        return ContentType.Draft;
      }
    }

    // Default to Alert
    return ContentType.Alert;
  }

  /**
   * Extract created date from SharePoint item
   */
  public static extractCreatedDate(item: any, fields: any): string {
    return item.createdDateTime ||
           fields?.CreatedDateTime ||
           fields?.Created ||
           "";
  }

  /**
   * Extract created by from SharePoint item
   */
  public static extractCreatedBy(item: any, fields: any): string {
    return item.createdBy?.user?.displayName ||
           fields?.CreatedBy?.LookupValue ||
           fields?.Author?.LookupValue ||
           item.author?.Title ||
           "Unknown";
  }

  /**
   * Safely parse metadata field
   */
  public static parseMetadata(value: any): any {
    if (!value) {
      return undefined;
    }

    if (typeof value !== 'string') {
      return value;
    }

    const parsed = JsonUtils.safeParse(value);
    if (parsed === null) {
      logger.warn('AlertTransformers', 'Failed to parse alert metadata; ignoring value', {
        valueSnippet: value.slice(0, 100)
      });
    }
    return parsed || undefined;
  }

  /**
   * Create original list item for multi-language support
   * Used by SharePointAlertService for complete data preservation
   */
  public static createOriginalListItem(item: any, fields: any): IAlertListItem {
    return {
      Id: parseInt(item.id.toString()),
      Title: fields.Title || '',
      Description: fields.Description || '',
      AlertType: fields.AlertType?.LookupValue || fields.AlertType || '',
      Priority: fields.Priority || AlertPriority.Medium,
      IsPinned: fields.IsPinned || false,
      NotificationType: fields.NotificationType || NotificationType.None,
      LinkUrl: fields.LinkUrl || '',
      LinkDescription: fields.LinkDescription || '',
      TargetSites: fields.TargetSites || '',
      Status: fields.Status || 'Active',
      Created: fields.Created || item.createdDateTime,
      Author: {
        Title: item.createdBy?.user?.displayName || item.author?.Title || 'Unknown'
      },
      ScheduledStart: fields.ScheduledStart || undefined,
      ScheduledEnd: fields.ScheduledEnd || undefined,
      Metadata: fields.Metadata || undefined,

      // Add all multi-language fields
      Title_EN: fields.Title_EN || '',
      Title_FR: fields.Title_FR || '',
      Title_DE: fields.Title_DE || '',
      Title_ES: fields.Title_ES || '',
      Title_SV: fields.Title_SV || '',
      Title_FI: fields.Title_FI || '',
      Title_DA: fields.Title_DA || '',
      Title_NO: fields.Title_NO || '',

      Description_EN: fields.Description_EN || '',
      Description_FR: fields.Description_FR || '',
      Description_DE: fields.Description_DE || '',
      Description_ES: fields.Description_ES || '',
      Description_SV: fields.Description_SV || '',
      Description_FI: fields.Description_FI || '',
      Description_DA: fields.Description_DA || '',
      Description_NO: fields.Description_NO || '',

      LinkDescription_EN: fields.LinkDescription_EN || '',
      LinkDescription_FR: fields.LinkDescription_FR || '',
      LinkDescription_DE: fields.LinkDescription_DE || '',
      LinkDescription_ES: fields.LinkDescription_ES || '',
      LinkDescription_SV: fields.LinkDescription_SV || '',
      LinkDescription_FI: fields.LinkDescription_FI || '',
      LinkDescription_DA: fields.LinkDescription_DA || '',
      LinkDescription_NO: fields.LinkDescription_NO || '',

      // Language and classification properties
      ItemType: fields.ItemType || '',
      TargetLanguage: fields.TargetLanguage || '',
      LanguageGroup: fields.LanguageGroup || '',
      AvailableForAll: fields.AvailableForAll || false,

      // Include any additional dynamic language fields
      ...Object.keys(fields)
        .filter(key => key.match(/^(Title|Description|LinkDescription)_[A-Z]{2}$/))
        .reduce((acc, key) => ({ ...acc, [key]: fields[key] }), {})
    };
  }

  /**
   * Map SharePoint item to IAlertItem (simplified version for AlertsContext)
   * @param item - SharePoint list item from Graph API
   * @param siteId - Site ID for alert identification
   * @param includeOriginalItem - Whether to include _originalListItem (default: false)
   */
  public static mapSharePointItemToAlert(
    item: any,
    siteId: string,
    includeOriginalItem: boolean = false
  ): IAlertItem {
    const fields = item.fields;

    // Parse all the fields using helper methods
    const priority = this.parsePriority(fields.Priority);
    const targetUsers = this.mapTargetUsers(fields.TargetUsers);
    const contentType = this.parseContentType(fields);
    const targetSites = this.parseTargetSitesField(fields.TargetSites);
    const createdDate = this.extractCreatedDate(item, fields);
    const createdBy = this.extractCreatedBy(item, fields);

    const alert: IAlertItem = {
      id: `${siteId}-${item.id}`,
      title: fields.Title || "",
      description: fields.Description || "",
      AlertType: fields.AlertType?.LookupValue || fields.AlertType || "Default",
      priority: priority,
      isPinned: fields.IsPinned || false,
      targetUsers: targetUsers,
      notificationType: fields.NotificationType || NotificationType.None,
      linkUrl: fields.LinkUrl || "",
      linkDescription: fields.LinkDescription || "Learn More",
      targetSites: targetSites,
      status: fields.Status || 'Active',
      createdDate,
      createdBy,
      contentType: contentType,
      targetLanguage: (fields.TargetLanguage as TargetLanguage) || TargetLanguage.All,
      languageGroup: fields.LanguageGroup || undefined,
      availableForAll: typeof fields.AvailableForAll === 'boolean'
        ? fields.AvailableForAll
        : Boolean(fields.AvailableForAll),
      scheduledStart: fields.ScheduledStart || undefined,
      scheduledEnd: fields.ScheduledEnd || undefined,
      metadata: this.parseMetadata(fields.Metadata),
      attachments: fields.AttachmentFiles?.map((file: any) => ({
        fileName: file.FileName || file.fileName || '',
        serverRelativeUrl: file.ServerRelativeUrl || file.serverRelativeUrl || '',
        size: file.Length || file.length || undefined
      })) || []
    };

    // Add original list item if requested (for SharePointAlertService)
    if (includeOriginalItem) {
      alert._originalListItem = this.createOriginalListItem(item, fields);
    }

    return alert;
  }
}
