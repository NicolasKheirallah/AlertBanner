import {
  IAlertItem,
  IAlertListItem,
  AlertPriority,
  NotificationType,
  ContentType,
  TargetLanguage,
  IPersonField,
  TranslationStatus,
  ContentStatus,
  IGraphListItem,
} from "../Alerts/IAlerts";
import { logger } from "../Services/LoggerService";
import { JsonUtils } from "./JsonUtils";
import { SiteIdUtils } from "./SiteIdUtils";
import { DEFAULT_ALERT_TYPE_NAME } from "./AppConstants";

export class AlertTransformers {
  // Parse target sites field from SharePoint - handles arrays, JSON strings, and CSV strings
  public static parseTargetSitesField(raw: any): string[] {
    if (!raw) {
      return [];
    }

    if (Array.isArray(raw)) {
      return raw.map(String).filter(Boolean);
    }

    if (typeof raw === "string") {
      const trimmed = raw.trim();
      if (!trimmed) {
        return [];
      }

      const parsed = JsonUtils.safeParse<string[]>(trimmed);
      if (parsed && Array.isArray(parsed)) {
        return parsed.map((entry) => String(entry)).filter(Boolean);
      }

      return trimmed
        .split(",")
        .map((site) => site.trim())
        .filter(Boolean);
    }

    return [];
  }

  public static mapPersonFieldData(
    personField: any,
    isGroup: boolean,
  ): IPersonField {
    if (personField.LookupId && personField.LookupValue) {
      return {
        id: personField.LookupId.toString(),
        displayName: personField.LookupValue,
        isGroup: isGroup,
      };
    }

    if (personField.ID || personField.id) {
      return {
        id: (personField.ID || personField.id).toString(),
        displayName: personField.Title || personField.displayName || "",
        email: personField.EMail || personField.email,
        loginName: personField.Name || personField.loginName,
        isGroup: isGroup,
      };
    }

    return {
      id: personField.id?.toString() || "",
      displayName: personField.displayName || personField.title || "",
      email: personField.email || personField.mail || "",
      loginName: personField.loginName || personField.userPrincipalName || "",
      isGroup: isGroup,
    };
  }

  // Map target users from SharePoint field - handles both single user and array of users
  public static mapTargetUsers(targetUsersField: any): IPersonField[] {
    if (!targetUsersField) {
      return [];
    }

    try {
      if (Array.isArray(targetUsersField)) {
        return targetUsersField.map((user: any) =>
          this.mapPersonFieldData(user, user.isGroup || false),
        );
      } else {
        return [
          this.mapPersonFieldData(
            targetUsersField,
            targetUsersField.isGroup || false,
          ),
        ];
      }
    } catch (error) {
      logger.warn("AlertTransformers", "Error processing target users", error);
      return [];
    }
  }

  public static parsePriority(priorityField: any): AlertPriority {
    if (!priorityField) {
      return AlertPriority.Medium;
    }

    try {
      const normalizedValue = String(priorityField).toLowerCase().trim();

      if (normalizedValue === AlertPriority.Critical)
        return AlertPriority.Critical;
      if (normalizedValue === AlertPriority.High) return AlertPriority.High;
      if (normalizedValue === AlertPriority.Medium) return AlertPriority.Medium;
      if (normalizedValue === AlertPriority.Low) return AlertPriority.Low;

      const priority =
        AlertPriority[priorityField as keyof typeof AlertPriority];
      return priority || AlertPriority.Medium;
    } catch {
      return AlertPriority.Medium;
    }
  }

  // Determine content type from SharePoint fields - checks both ItemType and ContentType fields
  public static parseContentType(fields: any): ContentType {
    const status = (fields.Status || "").toLowerCase();

    if (fields.ItemType) {
      const itemType = fields.ItemType.toLowerCase();
      if (itemType === "template") {
        return ContentType.Template;
      } else if (itemType === "draft") {
        return ContentType.Draft;
      } else if (itemType === "alert") {
        return status === "draft" ? ContentType.Draft : ContentType.Alert;
      }
    }

    if (fields.ContentType) {
      const contentType = fields.ContentType.toLowerCase();
      if (contentType === "template") {
        return ContentType.Template;
      } else if (contentType === "draft") {
        return ContentType.Draft;
      }
    }

    if (status === "draft") {
      return ContentType.Draft;
    }

    return ContentType.Alert;
  }

  private static extractCreatedDate(item: any, fields: any): string {
    return (
      item.createdDateTime || fields?.CreatedDateTime || fields?.Created || ""
    );
  }

  private static extractCreatedBy(item: any, fields: any): string {
    return (
      item.createdBy?.user?.displayName ||
      fields?.CreatedBy?.LookupValue ||
      fields?.Author?.LookupValue ||
      item.author?.Title ||
      "Unknown"
    );
  }

  private static parseMetadata(value: any): any {
    if (!value) {
      return undefined;
    }

    if (typeof value !== "string") {
      return value;
    }

    const parsed = JsonUtils.safeParse(value);
    if (parsed === null) {
      logger.warn(
        "AlertTransformers",
        "Failed to parse alert metadata; ignoring value",
        {
          valueSnippet: value.slice(0, 100),
        },
      );
    }
    return parsed || undefined;
  }

  private static createOriginalListItem(
    item: any,
    fields: any,
  ): IAlertListItem {
    return {
      Id: parseInt(item.id.toString()),
      Title: fields.Title || "",
      Description: fields.Description || "",
      AlertType: fields.AlertType?.LookupValue || fields.AlertType || "",
      Priority: fields.Priority || AlertPriority.Medium,
      IsPinned: fields.IsPinned || false,
      NotificationType: fields.NotificationType || NotificationType.None,
      LinkUrl: fields.LinkUrl || "",
      LinkDescription: fields.LinkDescription || "",
      TargetSites: fields.TargetSites || "",
      Status: fields.Status || "Active",
      Created: fields.Created || item.createdDateTime,
      Author: {
        Title:
          item.createdBy?.user?.displayName || item.author?.Title || "Unknown",
      },
      ScheduledStart: fields.ScheduledStart || undefined,
      ScheduledEnd: fields.ScheduledEnd || undefined,
      Metadata: fields.Metadata || undefined,

      ItemType: fields.ItemType || "",
      TargetLanguage: fields.TargetLanguage || "",
      LanguageGroup: fields.LanguageGroup || "",
      AvailableForAll: fields.AvailableForAll || false,
      TranslationStatus: fields.TranslationStatus || TranslationStatus.Approved,
      ContentStatus: fields.ContentStatus || ContentStatus.Approved,
      Reviewer: fields.Reviewer || undefined,
      ReviewNotes: fields.ReviewNotes || "",
      SubmittedDate: fields.SubmittedDate || undefined,
      ReviewedDate: fields.ReviewedDate || undefined,
    };
  }

  // Normalize site ID to extract the site GUID for consistent alert ID generation
  private static normalizeSiteIdForAlertId(siteId: string): string {
    if (!siteId) {
      return "";
    }

    if (siteId.includes(",")) {
      const extracted = SiteIdUtils.extractGuidFromGraphId(siteId);
      if (extracted) return extracted;
    }

    return SiteIdUtils.normalizeGuid(siteId);
  }

  public static mapSharePointItemToAlert(
    item: IGraphListItem,
    siteId: string,
    includeOriginalItem: boolean = false,
  ): IAlertItem {
    const fields = item.fields as any;

    const priority = this.parsePriority(fields.Priority);
    const targetUsers = this.mapTargetUsers(fields.TargetUsers);
    const contentType = this.parseContentType(fields);
    const targetSites = this.parseTargetSitesField(fields.TargetSites);
    const createdDate = this.extractCreatedDate(item, fields);
    const createdBy = this.extractCreatedBy(item, fields);

    const reviewer = fields.Reviewer
      ? this.mapTargetUsers(fields.Reviewer)
      : undefined;

    // Normalize site ID to ensure consistent alert IDs regardless of input format
    const normalizedSiteId = this.normalizeSiteIdForAlertId(siteId);

    const alert: IAlertItem = {
      id: `${normalizedSiteId}-${item.id}`,
      title: fields.Title || "",
      description:
        fields.Description ||
        fields.description ||
        fields.Body ||
        fields.body ||
        "",
      AlertType:
        fields.AlertType?.LookupValue ||
        fields.AlertType ||
        DEFAULT_ALERT_TYPE_NAME,
      priority: priority,
      isPinned: fields.IsPinned || false,
      targetUsers: targetUsers,
      notificationType: fields.NotificationType || NotificationType.None,
      linkUrl: fields.LinkUrl || "",
      linkDescription: fields.LinkDescription || "Learn More",
      targetSites: targetSites,
      status:
        (fields.Status as "Active" | "Expired" | "Scheduled" | "Draft") ||
        "Active",
      createdDate,
      createdBy,
      contentType: contentType,
      modified:
        item.lastModifiedDateTime || (fields.Modified as string | undefined),
      targetLanguage:
        (fields.TargetLanguage as TargetLanguage) || TargetLanguage.All,
      languageGroup: (fields.LanguageGroup as string) || undefined,
      availableForAll:
        typeof fields.AvailableForAll === "boolean"
          ? fields.AvailableForAll
          : Boolean(fields.AvailableForAll),
      translationStatus:
        (fields.TranslationStatus as TranslationStatus) ||
        TranslationStatus.Approved,
      sortOrder:
        typeof fields.SortOrder === "number" ? fields.SortOrder : undefined,
      contentStatus:
        (fields.ContentStatus as ContentStatus) || ContentStatus.Draft,
      reviewer: reviewer,
      reviewNotes: (fields.ReviewNotes as string) || "",
      submittedDate: (fields.SubmittedDate as string) || undefined,
      reviewedDate: (fields.ReviewedDate as string) || undefined,
      scheduledStart: (fields.ScheduledStart as string) || undefined,
      scheduledEnd: (fields.ScheduledEnd as string) || undefined,
      metadata: this.parseMetadata(fields.Metadata),
      attachments:
        (fields.AttachmentFiles as any[])?.map((file: any) => ({
          fileName: file.FileName || file.fileName || "",
          serverRelativeUrl:
            file.ServerRelativeUrl || file.serverRelativeUrl || "",
          size: file.Length || file.length || undefined,
        })) || [],
    };

    if (includeOriginalItem) {
      alert._originalListItem = this.createOriginalListItem(item, fields);
    }

    return alert;
  }
}
