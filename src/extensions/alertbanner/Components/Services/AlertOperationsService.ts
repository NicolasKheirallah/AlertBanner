import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SharePointListLocator } from "./SharePointListLocator";
import { AlertTransformers } from "../Utils/AlertTransformers";
import { AlertFilters } from "../Utils/AlertFilters";
import { logger } from "./LoggerService";
import { ErrorUtils } from "../Utils/ErrorUtils";
import { JsonUtils } from "../Utils/JsonUtils";
import { StringUtils } from "../Utils/StringUtils";
import {
  LIST_NAMES,
  SUPPORTED_LANGUAGES,
  DEFAULT_ALERT_TYPES,
  API_CONFIG,
  ALERT_ITEM_TYPES,
} from "../Utils/AppConstants";
import {
  IAlertItem,
  IAlertType,
  ContentType,
  AlertPriority,
  ContentStatus,
} from "../Alerts/IAlerts";
import { SPHttpClient } from "@microsoft/sp-http";
import { RetryUtils } from "../Utils/RetryUtils";
import { PermissionService } from "./PermissionService";
import {
  DEFAULT_LANGUAGE_POLICY,
  ILanguagePolicy,
  normalizeLanguagePolicy,
} from "./LanguagePolicyService";

export class AlertOperationsService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private locator: SharePointListLocator;
  private permissionService: PermissionService;
  private validatedListSchemas: Set<string> = new Set();
  private readonly maxItemsToFetch: number = API_CONFIG.MAX_PAGE_SIZE * 5;
  private readonly languagePolicyTitle = "LanguagePolicy";

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
    locator: SharePointListLocator,
  ) {
    this.graphClient = graphClient;
    this.context = context;
    this.locator = locator;
    this.permissionService = PermissionService.getInstance(context);
  }

  public async getActiveAlerts(siteId: string): Promise<IAlertItem[]> {
    const dateTimeNow = new Date().toISOString();
    const alertsListApi = await this.locator.getAlertsListApi(siteId);

    if (!alertsListApi) {
      logger.info(
        "AlertOperationsService",
        `Alerts list not found for site ${siteId}, returning empty results`,
      );
      return [];
    }

    const availableColumns =
      await this.locator.getAvailableColumns(alertsListApi);

    const filterParts: string[] = [];
    if (availableColumns.has("ScheduledStart")) {
      filterParts.push(
        `(fields/ScheduledStart le '${dateTimeNow}' or fields/ScheduledStart eq null)`,
      );
    }
    if (availableColumns.has("ScheduledEnd")) {
      filterParts.push(
        `(fields/ScheduledEnd ge '${dateTimeNow}' or fields/ScheduledEnd eq null)`,
      );
    }
    const filterQuery =
      filterParts.length > 0 ? filterParts.join(" and ") : undefined;
    const baseFields = [
      "Title",
      "AlertType",
      "Description",
      "ScheduledStart",
      "ScheduledEnd",
      "Priority",
      "IsPinned",
      "NotificationType",
      "LinkUrl",
      "LinkDescription",
      "TargetSites",
      "Status",
      "ItemType",
      "TargetLanguage",
      "LanguageGroup",
      "Attachments",
    ];
    const optionalFields = ["AvailableForAll", "Metadata"];
    if (availableColumns.has("TranslationStatus")) {
      optionalFields.push("TranslationStatus");
    }
    // Approval Workflow
    if (availableColumns.has("ContentStatus")) {
      optionalFields.push("ContentStatus");
      optionalFields.push("Reviewer");
      optionalFields.push("ReviewNotes");
      optionalFields.push("SubmittedDate");
      optionalFields.push("ReviewedDate");
    }

    const selectedFields = [
      ...baseFields.filter(
        (f) => availableColumns.has(f) || ["Title", "Attachments"].includes(f),
      ),
      ...optionalFields.filter((f) => availableColumns.has(f)),
    ];
    if (!selectedFields.includes("Title")) selectedFields.unshift("Title");

    const executeQuery = async (listApi: string): Promise<IAlertItem[]> => {
      let request = this.graphClient
        .api(`${listApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand(`fields($select=${selectedFields.join(",")})`)
        .orderby(
          availableColumns.has("ScheduledStart")
            ? "fields/ScheduledStart desc"
            : "fields/Created desc",
        )
        .top(API_CONFIG.MAX_PAGE_SIZE);

      if (availableColumns.has("ContentStatus")) {
        // Only show Approved alerts in the main feed
        request = request.filter(
          filterQuery
            ? `${filterQuery} and fields/ContentStatus eq 'Approved'`
            : "fields/ContentStatus eq 'Approved'",
        );
      } else if (filterQuery) {
        request = request.filter(filterQuery);
      }

      const items = await this.fetchPagedItems(request);
      const resolvedSiteId =
        await this.locator.ensureGraphSiteIdentifier(siteId);
      const listId = await this.locator.resolveListId(
        siteId,
        LIST_NAMES.ALERTS,
      );

      let alerts = items.map((item: any) =>
        AlertTransformers.mapSharePointItemToAlert(item, resolvedSiteId),
      );
      alerts = await this.attachAttachments(siteId, listId, items, alerts);
      return AlertFilters.excludeNonPublicAlerts(alerts);
    };

    try {
      return await executeQuery(alertsListApi);
    } catch (error) {
      const statusCode =
        (error as any)?.statusCode ||
        (error as any)?.status ||
        (error as any)?.code;
      const errorMessage = (error as any)?.message || "";
      const isNotFound =
        statusCode === 404 ||
        statusCode === "itemNotFound" ||
        errorMessage.toLowerCase().includes("not found");

      if (isNotFound) {
        logger.warn(
          "AlertOperationsService",
          `Alerts list not found for site ${siteId}. Invalidating cache and retrying.`,
          { siteId },
        );
        await this.locator.invalidateListId(siteId, LIST_NAMES.ALERTS);

        try {
          const refreshedListApi = await this.locator.getAlertsListApi(siteId);
          if (!refreshedListApi) {
            return [];
          }
          return await executeQuery(refreshedListApi);
        } catch (retryError) {
          logger.warn(
            "AlertOperationsService",
            `Retry failed after invalidating cache for site ${siteId}.`,
            retryError,
          );
          return [];
        }
      }

      logger.error(
        "AlertOperationsService",
        `Failed to get active alerts from ${siteId}`,
        error,
      );
      return [];
    }
  }

  public async createAlert(
    alert: Omit<IAlertItem, "id" | "createdDate" | "createdBy" | "status"> &
      Partial<Pick<IAlertItem, "status">>,
  ): Promise<IAlertItem> {
    const siteId = this.context.pageContext.site.id.toString();

    return this.permissionService.executeWriteOperation(
      async () => {
        if (!alert.title?.trim()) throw new Error("Alert title is required");
        if (!alert.description?.trim())
          throw new Error("Alert description is required");
        if (!alert.AlertType?.trim()) throw new Error("Alert type is required");
        if (!alert.targetSites || alert.targetSites.length === 0)
          throw new Error("At least one target site is required");

        const alertsListApi = await this.locator.getAlertsListApi(siteId);

        if (!alertsListApi) {
          throw new Error(
            `Alerts list not found for site ${siteId}. Please create the list first.`,
          );
        }

        const availableColumns =
          await this.locator.getAvailableColumns(alertsListApi);
        const required = [
          "Title",
          "Description",
          "AlertType",
          "Priority",
          "IsPinned",
        ];
        const missing = required.filter((c) => !availableColumns.has(c));
        if (missing.length > 0 && !availableColumns.has("Title")) {
        }

        const fields: any = {
          Title: alert.title.trim(),
          Description: alert.description.trim(),
          AlertType: alert.AlertType.trim(),
          Priority: alert.priority,
          IsPinned: Boolean(alert.isPinned),
          NotificationType: alert.notificationType,
        };

        if (alert.linkUrl?.trim()) fields.LinkUrl = alert.linkUrl.trim();
        if (alert.linkDescription?.trim())
          fields.LinkDescription = alert.linkDescription.trim();
        if (alert.targetSites?.length > 0)
          fields.TargetSites = JsonUtils.safeStringify(alert.targetSites);

        fields.Status =
          alert.status ||
          (alert.scheduledStart && new Date(alert.scheduledStart) > new Date()
            ? "Scheduled"
            : "Active");
        if (alert.scheduledStart)
          fields.ScheduledStart = new Date(alert.scheduledStart).toISOString();
        if (alert.scheduledEnd)
          fields.ScheduledEnd = new Date(alert.scheduledEnd).toISOString();
        if (alert.metadata)
          fields.Metadata = JsonUtils.safeStringify(alert.metadata);
        if ((alert.targetUsers?.length || 0) > 0)
          fields.TargetUsers = alert.targetUsers;

        fields.ItemType = alert.contentType;
        fields.TargetLanguage = alert.targetLanguage;
        if (alert.languageGroup) fields.LanguageGroup = alert.languageGroup;
        fields.AvailableForAll = Boolean(alert.availableForAll);
        if (
          alert.translationStatus &&
          availableColumns.has("TranslationStatus")
        ) {
          fields.TranslationStatus = alert.translationStatus;
        }

        // Approval Workflow
        if (availableColumns.has("ContentStatus")) {
          fields.ContentStatus = alert.contentStatus || ContentStatus.Draft;
          if (alert.reviewer) fields.Reviewer = alert.reviewer;
          if (alert.reviewNotes) fields.ReviewNotes = alert.reviewNotes;
        }

        let response = await this.graphClient
          .api(`${alertsListApi}/items`)
          .post({ fields });

        const created = await this.graphClient
          .api(`${alertsListApi}/items/${response.id}`)
          .expand("fields")
          .get();
        return AlertTransformers.mapSharePointItemToAlert(created, siteId);
      },
      {
        operation: "CREATE_ALERT",
        targetSite: siteId,
        targetList: LIST_NAMES.ALERTS,
        justification: `Creating ${alert.contentType || "alert"}: ${alert.title}`,
      },
    );
  }

  public async updateAlert(
    alertId: string,
    updates: Partial<IAlertItem>,
  ): Promise<IAlertItem> {
    const { siteId, itemId } = this.parseAlertId(alertId);

    return this.permissionService.executeWriteOperation(
      async () => {
        const alertsListApi = await this.locator.getAlertsListApi(siteId);

        if (!alertsListApi) {
          throw new Error(
            `Alerts list not found for site ${siteId}. Please create the list first.`,
          );
        }

        const fields: any = {};
        if (updates.title) fields.Title = updates.title;
        if (updates.description) fields.Description = updates.description;
        if (updates.AlertType) fields.AlertType = updates.AlertType;
        if (updates.priority) fields.Priority = updates.priority;
        if (updates.isPinned !== undefined) fields.IsPinned = updates.isPinned;
        if (updates.notificationType)
          fields.NotificationType = updates.notificationType;
        if (updates.linkUrl !== undefined) fields.LinkUrl = updates.linkUrl;
        if (updates.linkDescription !== undefined)
          fields.LinkDescription = updates.linkDescription;
        if (updates.targetSites)
          fields.TargetSites = JsonUtils.safeStringify(updates.targetSites);
        if (updates.scheduledStart !== undefined)
          fields.ScheduledStart = updates.scheduledStart;
        if (updates.scheduledEnd !== undefined)
          fields.ScheduledEnd = updates.scheduledEnd;
        if (updates.targetUsers !== undefined)
          fields.TargetUsers = updates.targetUsers;
        if (updates.metadata)
          fields.Metadata = JsonUtils.safeStringify(updates.metadata);
        if (updates.status) fields.Status = updates.status;
        if (updates.translationStatus) {
          try {
            const availableColumns =
              await this.locator.getAvailableColumns(alertsListApi);
            if (availableColumns.has("TranslationStatus")) {
              fields.TranslationStatus = updates.translationStatus;
            }
          } catch (error) {
            logger.warn(
              "AlertOperationsService",
              "Unable to verify TranslationStatus column; skipping update",
              error,
            );
          }
        }

        // Approval Workflow
        if (updates.contentStatus) fields.ContentStatus = updates.contentStatus;
        if (updates.reviewer !== undefined) fields.Reviewer = updates.reviewer;
        if (updates.reviewNotes !== undefined)
          fields.ReviewNotes = updates.reviewNotes;
        if (updates.submittedDate !== undefined)
          fields.SubmittedDate = updates.submittedDate;
        if (updates.reviewedDate !== undefined)
          fields.ReviewedDate = updates.reviewedDate;

        await this.graphClient
          .api(`${alertsListApi}/items/${itemId}/fields`)
          .patch(fields);

        const updated = await this.graphClient
          .api(`${alertsListApi}/items/${itemId}`)
          .expand("fields")
          .get();
        return AlertTransformers.mapSharePointItemToAlert(updated, siteId);
      },
      {
        operation: "UPDATE_ALERT",
        targetSite: siteId,
        targetList: LIST_NAMES.ALERTS,
        justification: `Updating alert ${itemId}`,
      },
    );
  }

  public async deleteAlert(alertId: string): Promise<void> {
    const { siteId, itemId } = this.parseAlertId(alertId);

    return this.permissionService.executeWriteOperation(
      async () => {
        const alertsListApi = await this.locator.getAlertsListApi(siteId);

        if (!alertsListApi) {
          throw new Error(
            `Alerts list not found for site ${siteId}. Please create the list first.`,
          );
        }

        await this.graphClient.api(`${alertsListApi}/items/${itemId}`).delete();
      },
      {
        operation: "DELETE_ALERT",
        targetSite: siteId,
        targetList: LIST_NAMES.ALERTS,
        justification: `Deleting alert item ${itemId}`,
      },
    );
  }

  public async deleteAlerts(alertIds: string[]): Promise<void> {
    await Promise.allSettled(alertIds.map((id) => this.deleteAlert(id)));
  }

  public parseAlertId(alertId: string): { siteId: string; itemId: string } {
    const lastHyphen = alertId.lastIndexOf("-");
    if (lastHyphen > 0 && lastHyphen < alertId.length - 1) {
      const site = alertId.substring(0, lastHyphen);
      const item = alertId.substring(lastHyphen + 1);
      if (/^\d+$/.test(item)) return { siteId: site, itemId: item };
    }
    return {
      siteId: this.context.pageContext.site.id.toString(),
      itemId: alertId,
    };
  }

  public async getAlertTypes(siteIdOverride?: string): Promise<IAlertType[]> {
    const siteId =
      siteIdOverride || this.context.pageContext.site.id.toString();
    try {
      const listApi = await this.locator.getAlertTypesListApi(siteId);
      const response = await this.graphClient
        .api(`${listApi}/items`)
        .expand("fields")
        .orderby("fields/SortOrder")
        .get();

      if (!response.value || response.value.length === 0)
        return DEFAULT_ALERT_TYPES.map((t) => ({ ...t }));

      return response.value.map((item: any) => {
        const parsedStyles =
          JsonUtils.safeParse(item.fields.PriorityStyles) || {};
        const defaultPriority = parsedStyles.__defaultPriority;

        return {
          name: item.fields.Title || "",
          iconName: item.fields.IconName || "Info",
          backgroundColor: item.fields.BackgroundColor || "#0078d4",
          textColor: item.fields.TextColor || "#ffffff",
          additionalStyles: item.fields.AdditionalStyles || "",
          priorityStyles: parsedStyles,
          defaultPriority: defaultPriority as AlertPriority,
        };
      });
    } catch (e) {
      return DEFAULT_ALERT_TYPES.map((t) => ({ ...t }));
    }
  }

  public async getDraftAlerts(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.locator.getAlertsListApi(siteId);

      if (!alertsListApi) {
        return [];
      }

      const currentUserEmail =
        this.context.pageContext.user.email ||
        (this.context.pageContext.user as any)?.userPrincipalName ||
        "";
      const availableColumns =
        await this.locator.getAvailableColumns(alertsListApi);

      const baseFields = [
        "Title",
        "AlertType",
        "Description",
        "Priority",
        "IsPinned",
        "NotificationType",
        "LinkUrl",
        "LinkDescription",
        "TargetSites",
        "Status",
        "ItemType",
        "TargetLanguage",
        "LanguageGroup",
        "ScheduledStart",
        "ScheduledEnd",
        "TargetUsers",
        "Author",
        "Modified",
      ];
      const selectedFields = baseFields.filter(
        (f) => availableColumns.has(f) || ["Title"].includes(f),
      );

      const filters: string[] = [];
      if (availableColumns.has("ItemType")) {
        filters.push("fields/ItemType eq 'draft'");
      } else if (availableColumns.has("Status")) {
        filters.push("tolower(fields/Status) eq 'draft'");
      }
      if (availableColumns.has("Author") && currentUserEmail) {
        const safeEmail = currentUserEmail.replace(/'/g, "''");
        filters.push(`fields/Author/Email eq '${safeEmail}'`);
      }

      let request = this.graphClient
        .api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand(`fields($select=${selectedFields.join(",")})`)
        .top(API_CONFIG.MAX_PAGE_SIZE);

      if (filters.length > 0) request = request.filter(filters.join(" and "));
      request = request.orderby(
        availableColumns.has("Modified")
          ? "fields/Modified desc"
          : "fields/Title asc",
      );

      const items = await this.fetchPagedItems(request);
      const listId = await this.locator.resolveListId(
        siteId,
        LIST_NAMES.ALERTS,
      );
      let alerts = items.map((item: any) =>
        AlertTransformers.mapSharePointItemToAlert(item, siteId),
      );
      alerts = await this.attachAttachments(siteId, listId, items, alerts);
      return alerts;
    } catch (e) {
      logger.warn("AlertOperationsService", "Failed to get drafts", e);
      return [];
    }
  }

  public async saveDraft(draft: Partial<IAlertItem>): Promise<IAlertItem> {
    const data: any = {
      ...draft,
      contentType: ContentType.Draft,
      status: "Draft",
    };
    if (draft.id && parseInt(draft.id) > 0)
      return this.updateAlert(draft.id, data);
    return this.createAlert(data);
  }

  public async deleteDraft(draftId: string): Promise<void> {
    return this.deleteAlert(draftId);
  }

  public async addAttachment(
    listId: string,
    itemId: number,
    fileName: string,
    fileContent: ArrayBuffer,
    siteId?: string,
  ): Promise<{ fileName: string; serverRelativeUrl: string }> {
    try {
      const siteUrl = siteId
        ? await this.locator.getSiteUrlFromIdentifier(siteId)
        : this.context.pageContext.web.absoluteUrl;
      const uploadUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`;
      const response = await this.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/octet-stream",
          },
          body: fileContent,
        },
      );
      if (!response.ok)
        throw new Error(`Upload failed: ${response.statusText}`);
      const result = await response.json();
      return {
        fileName: result.d.FileName,
        serverRelativeUrl: result.d.ServerRelativeUrl,
      };
    } catch (e) {
      logger.error("AlertOperationsService", "Upload attachment failed", e);
      throw e;
    }
  }

  public async deleteAttachment(
    listId: string,
    itemId: number,
    fileName: string,
    siteId?: string,
  ): Promise<void> {
    try {
      const siteUrl = siteId
        ? await this.locator.getSiteUrlFromIdentifier(siteId)
        : this.context.pageContext.web.absoluteUrl;
      const deleteUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
      await this.context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*",
          },
        },
      );
    } catch (e) {
      logger.error("AlertOperationsService", "Delete attachment failed", e);
      throw e;
    }
  }

  public async getAlertsForSite(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.locator.getAlertsListApi(siteId);
      let request = this.graphClient
        .api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand("fields")
        .top(API_CONFIG.MAX_PAGE_SIZE)
        .orderby("fields/Created desc");

      const items = await this.fetchPagedItems(request);
      const listId = await this.locator.resolveListId(
        siteId,
        LIST_NAMES.ALERTS,
      );
      const filteredItems = items.filter((item: any) => {
        const type = (item.fields?.ItemType || "").toLowerCase();
        const status = (item.fields?.Status || "").toLowerCase();
        const title = item.fields?.Title || "";
        return (
          type !== "draft" &&
          type !== "template" &&
          type !== ALERT_ITEM_TYPES.SETTINGS &&
          status !== "draft" &&
          !title.startsWith("[Auto-saved]") &&
          !title.startsWith("[auto-saved]")
        );
      });
      let alerts = filteredItems.map((item: any) =>
        AlertTransformers.mapSharePointItemToAlert(item, siteId),
      );
      alerts = await this.attachAttachments(
        siteId,
        listId,
        filteredItems,
        alerts,
      );
      return alerts;
    } catch (e) {
      logger.warn(
        "AlertOperationsService",
        `Failed to get alerts for site ${siteId}`,
        e,
      );
      return [];
    }
  }

  public async getTemplateAlerts(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.locator.getAlertsListApi(siteId);
      let request = this.graphClient
        .api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .filter("fields/ItemType eq 'template'")
        .expand("fields")
        .top(API_CONFIG.MAX_PAGE_SIZE);

      const items = await this.fetchPagedItems(request);
      const listId = await this.locator.resolveListId(
        siteId,
        LIST_NAMES.ALERTS,
      );
      let alerts = items.map((item: any) =>
        AlertTransformers.mapSharePointItemToAlert(item, siteId),
      );
      alerts = await this.attachAttachments(siteId, listId, items, alerts);
      return alerts;
    } catch (e) {
      logger.warn(
        "AlertOperationsService",
        `Failed to get templates ${siteId}`,
        e,
      );
      return [];
    }
  }

  public async getLanguagePolicy(siteId?: string): Promise<ILanguagePolicy> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();
    try {
      const alertsListApi = await this.locator.getAlertsListApi(targetSiteId);
      if (!alertsListApi) {
        return DEFAULT_LANGUAGE_POLICY;
      }

      const availableColumns =
        await this.locator.getAvailableColumns(alertsListApi);
      if (
        !availableColumns.has("ItemType") ||
        !availableColumns.has("Metadata")
      ) {
        return DEFAULT_LANGUAGE_POLICY;
      }

      const response = await this.graphClient
        .api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand("fields($select=Title,Metadata,ItemType)")
        .filter(`fields/ItemType eq '${ALERT_ITEM_TYPES.SETTINGS}'`)
        .top(5)
        .get();

      const items = response?.value || [];
      const policyItem =
        items.find(
          (item: any) =>
            (item.fields?.Title || "").toLowerCase() ===
            this.languagePolicyTitle.toLowerCase(),
        ) || items[0];

      const raw = policyItem?.fields?.Metadata;
      const parsed = JsonUtils.safeParse(raw);
      return normalizeLanguagePolicy(parsed || {});
    } catch (e) {
      logger.warn(
        "AlertOperationsService",
        "Failed to load language policy",
        e,
      );
      return DEFAULT_LANGUAGE_POLICY;
    }
  }

  public async saveLanguagePolicy(
    policy: ILanguagePolicy,
    siteId?: string,
  ): Promise<void> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();

    return this.permissionService.executeWriteOperation(
      async () => {
        const alertsListApi = await this.locator.getAlertsListApi(targetSiteId);
        if (!alertsListApi) {
          throw new Error(
            `Alerts list not found for site ${targetSiteId}. Please create the list first.`,
          );
        }

        const availableColumns =
          await this.locator.getAvailableColumns(alertsListApi);
        const payloadMetadata =
          JsonUtils.safeStringify(policy) ||
          JsonUtils.safeStringify(DEFAULT_LANGUAGE_POLICY) ||
          "{}";

        const response = await this.graphClient
          .api(`${alertsListApi}/items`)
          .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
          .expand("fields($select=Title,ItemType)")
          .filter(`fields/ItemType eq '${ALERT_ITEM_TYPES.SETTINGS}'`)
          .top(5)
          .get();

        const items = response?.value || [];
        const existing =
          items.find(
            (item: any) =>
              (item.fields?.Title || "").toLowerCase() ===
              this.languagePolicyTitle.toLowerCase(),
          ) || items[0];

        const fields: any = {
          Title: this.languagePolicyTitle,
          Metadata: payloadMetadata,
          ItemType: ALERT_ITEM_TYPES.SETTINGS,
        };

        if (availableColumns.has("Status")) {
          fields.Status = "Active";
        }

        if (existing?.id) {
          await this.graphClient
            .api(`${alertsListApi}/items/${existing.id}/fields`)
            .patch(fields);
          return;
        }

        await this.graphClient.api(`${alertsListApi}/items`).post({ fields });
      },
      {
        operation: "UPDATE_LANGUAGE_POLICY",
        targetSite: targetSiteId,
        targetList: LIST_NAMES.ALERTS,
        justification: "Update language policy configuration",
      },
    );
  }

  public async updateAlertStatuses(siteIds?: string[]): Promise<void> {}

  public async saveAlertTypes(alertTypes: IAlertType[]): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const alertTypesListApi = await this.locator.getAlertTypesListApi(siteId);

      const existingItems = await this.graphClient
        .api(`${alertTypesListApi}/items`)
        .expand("fields")
        .get();
      for (const item of existingItems.value) {
        await this.graphClient
          .api(`${alertTypesListApi}/items/${item.id}`)
          .delete();
      }

      for (let i = 0; i < alertTypes.length; i++) {
        const alertType = alertTypes[i];
        await this.graphClient.api(`${alertTypesListApi}/items`).post({
          fields: {
            Title: alertType.name,
            IconName: alertType.iconName,
            BackgroundColor: alertType.backgroundColor,
            TextColor: alertType.textColor,
            AdditionalStyles: alertType.additionalStyles || "",
            PriorityStyles:
              JsonUtils.safeStringify({
                ...(alertType.priorityStyles || {}),
                __defaultPriority: alertType.defaultPriority || undefined,
              }) || "{}",
            SortOrder: i,
          },
        });
      }
    } catch (error) {
      logger.error(
        "AlertOperationsService",
        "Failed to save alert types",
        error,
      );
      throw error;
    }
  }

  private async fetchPagedItems(request: any): Promise<any[]> {
    const items: any[] = [];

    interface GraphResponse {
      value?: any[];
      "@odata.nextLink"?: string;
    }

    let response: GraphResponse = await RetryUtils.executeWithRetry(
      () => request.get(),
      {
        suppressFailureLog: (error) => ErrorUtils.isListNotFoundError(error),
      },
    );
    if (Array.isArray(response?.value)) {
      items.push(...response.value);
    }

    let nextLink = response?.["@odata.nextLink"];
    while (nextLink && items.length < this.maxItemsToFetch) {
      const nextRequest = this.graphClient.api(nextLink);
      response = await RetryUtils.executeWithRetry(() => nextRequest.get());
      if (Array.isArray(response?.value)) {
        items.push(...response.value);
      }
      nextLink = response?.["@odata.nextLink"];
    }

    return items;
  }

  private async attachAttachments(
    siteId: string,
    listId: string,
    items: any[],
    alerts: IAlertItem[],
  ): Promise<IAlertItem[]> {
    const attachmentsTasks = items.map(async (item: any, index: number) => {
      const hasAttachments = Boolean(item?.fields?.Attachments);
      if (!hasAttachments) {
        return;
      }

      try {
        const attachments = await this.getAttachmentsForItem(
          siteId,
          listId,
          item.id,
        );
        alerts[index].attachments = attachments;
      } catch (error) {
        logger.warn("AlertOperationsService", "Failed to fetch attachments", {
          itemId: item.id,
          siteId,
          error,
        });
      }
    });

    await Promise.allSettled(attachmentsTasks);
    return alerts;
  }

  private async getAttachmentsForItem(
    siteId: string,
    listId: string,
    itemId: number | string,
  ): Promise<{ fileName: string; serverRelativeUrl: string; size?: number }[]> {
    const siteUrl = await this.locator.getSiteUrlFromIdentifier(siteId);
    const safeItemId =
      typeof itemId === "string" ? parseInt(itemId, 10) : itemId;

    if (!safeItemId || Number.isNaN(safeItemId)) {
      return [];
    }

    const attachmentUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${safeItemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl,Length`;
    const response = await this.context.spHttpClient.get(
      attachmentUrl,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata" },
      },
    );

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    const files = data?.value || data?.d?.results || [];

    return files.map((file: any) => ({
      fileName: file.FileName || file.Name || "",
      serverRelativeUrl: file.ServerRelativeUrl || file.serverRelativeUrl,
      size: file.Length || file.length || undefined,
    }));
  }

  // Approval Workflow Methods
  public async submitAlert(
    alertId: string,
    reviewerId?: string,
  ): Promise<IAlertItem> {
    const updates: Partial<IAlertItem> = {
      contentStatus: ContentStatus.PendingReview,
      submittedDate: new Date().toISOString(),
    };
    // If reviewerId is provided, we would likely update the Reviewer field here
    // But Reviewer is a Person field, passing just ID might require lookup or specific format
    // For now, we'll assume the UI handles assignment or we just set status
    return this.updateAlert(alertId, updates);
  }

  public async approveAlert(
    alertId: string,
    comments?: string,
  ): Promise<IAlertItem> {
    const updates: Partial<IAlertItem> = {
      contentStatus: ContentStatus.Approved,
      reviewedDate: new Date().toISOString(),
      reviewNotes: comments,
    };
    return this.updateAlert(alertId, updates);
  }

  public async rejectAlert(
    alertId: string,
    comments?: string,
  ): Promise<IAlertItem> {
    const updates: Partial<IAlertItem> = {
      contentStatus: ContentStatus.Rejected,
      reviewedDate: new Date().toISOString(),
      reviewNotes: comments,
    };
    return this.updateAlert(alertId, updates);
  }
}
