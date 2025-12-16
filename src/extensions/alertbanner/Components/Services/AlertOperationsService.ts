import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SharePointListLocator } from "./SharePointListLocator";
import { AlertTransformers } from "../Utils/AlertTransformers";
import { AlertFilters } from "../Utils/AlertFilters";
import { logger } from "./LoggerService";
import { ErrorUtils } from "../Utils/ErrorUtils";
import { JsonUtils } from "../Utils/JsonUtils";
import { StringUtils } from "../Utils/StringUtils";
import { LIST_NAMES, SUPPORTED_LANGUAGES, DEFAULT_ALERT_TYPES } from "../Utils/AppConstants";
import { IAlertItem, IAlertType, ContentType, AlertPriority } from "../Alerts/IAlerts";
import { SPHttpClient } from "@microsoft/sp-http";

export class AlertOperationsService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private locator: SharePointListLocator;
  // We keep validatedListSchemas locally or inside Locator. 
  // For now, local set to avoid re-validating every time is fine, scoped to this service instance.
  private validatedListSchemas: Set<string> = new Set();

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
    locator: SharePointListLocator
  ) {
    this.graphClient = graphClient;
    this.context = context;
    this.locator = locator;
  }

  public async getActiveAlerts(siteId: string): Promise<IAlertItem[]> {
    const dateTimeNow = new Date().toISOString();
    const alertsListApi = await this.locator.getAlertsListApi(siteId);
    const availableColumns = await this.locator.getAvailableColumns(alertsListApi);

    const filterParts: string[] = [];
    if (availableColumns.has("ScheduledStart")) {
      filterParts.push(`(fields/ScheduledStart le '${dateTimeNow}' or fields/ScheduledStart eq null)`);
    }
    if (availableColumns.has("ScheduledEnd")) {
      filterParts.push(`(fields/ScheduledEnd ge '${dateTimeNow}' or fields/ScheduledEnd eq null)`);
    }
    if (availableColumns.has("ItemType")) {
      filterParts.push(`(fields/ItemType ne 'template')`);
      filterParts.push(`(fields/ItemType ne 'draft')`);
    }
    if (availableColumns.has("Status")) {
      filterParts.push(`(fields/Status ne 'draft')`);
    }

    const filterQuery = filterParts.length > 0 ? filterParts.join(" and ") : undefined;
    const baseFields = ["Title", "AlertType", "Description", "ScheduledStart", "ScheduledEnd", "Priority", "IsPinned", "NotificationType", "LinkUrl", "LinkDescription", "TargetSites", "Status", "ItemType", "TargetLanguage", "LanguageGroup", "Attachments"];
    const optionalFields = ["AvailableForAll", "Metadata"];
    
    const selectedFields = [
      ...baseFields.filter(f => availableColumns.has(f) || ["Title", "Attachments"].includes(f)),
      ...optionalFields.filter(f => availableColumns.has(f))
    ];
    if (!selectedFields.includes("Title")) selectedFields.unshift("Title");

    try {
      let request = this.graphClient.api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand(`fields($select=${selectedFields.join(",")})`)
        .orderby(availableColumns.has("ScheduledStart") ? "fields/ScheduledStart desc" : "fields/Created desc")
        .top(25);

      if (filterQuery) request = request.filter(filterQuery);

      const response = await request.get();
      const resolvedSiteId = await this.locator.ensureGraphSiteIdentifier(siteId);
      
      let alerts = response.value.map((item: any) => AlertTransformers.mapSharePointItemToAlert(item, resolvedSiteId));
      return AlertFilters.excludeNonPublicAlerts(alerts);
    } catch (error) {
      logger.error("AlertOperationsService", `Failed to get active alerts from ${siteId}`, error);
      return [];
    }
  }

  public async createAlert(alert: Omit<IAlertItem, "id" | "createdDate" | "createdBy" | "status"> & Partial<Pick<IAlertItem, "status">>): Promise<IAlertItem> {
      // Logic copied from Original, condensed
      try {
        const siteId = this.context.pageContext.site.id.toString();
        if (!alert.title?.trim()) throw new Error("Alert title is required");
        if (!alert.description?.trim()) throw new Error("Alert description is required");
        if (!alert.AlertType?.trim()) throw new Error("Alert type is required");
        if (!alert.targetSites || alert.targetSites.length === 0) throw new Error("At least one target site is required");

        const alertsListApi = await this.locator.getAlertsListApi(siteId);
        
        // Validation Logic (Schema check) - simplified to use availableColumns from Locator
        const availableColumns = await this.locator.getAvailableColumns(alertsListApi);
        const required = ["Title", "Description", "AlertType", "Priority", "IsPinned"];
        const missing = required.filter(c => !availableColumns.has(c));
        if (missing.length > 0 && !availableColumns.has("Title")) { // Locator fallback might have Title only. 
            // If fallback, we might not know real columns. Trust the post.
        }

        const fields: any = {
            Title: alert.title.trim(),
            Description: alert.description.trim(),
            AlertType: alert.AlertType.trim(),
            Priority: alert.priority,
            IsPinned: Boolean(alert.isPinned),
            NotificationType: alert.notificationType
        };

        if (alert.linkUrl?.trim()) fields.LinkUrl = alert.linkUrl.trim();
        if (alert.linkDescription?.trim()) fields.LinkDescription = alert.linkDescription.trim();
        if (alert.targetSites?.length > 0) fields.TargetSites = JsonUtils.safeStringify(alert.targetSites);
        
        fields.Status = alert.status || (alert.scheduledStart && new Date(alert.scheduledStart) > new Date() ? "Scheduled" : "Active");
        if (alert.scheduledStart) fields.ScheduledStart = new Date(alert.scheduledStart).toISOString();
        if (alert.scheduledEnd) fields.ScheduledEnd = new Date(alert.scheduledEnd).toISOString();
        if (alert.metadata) fields.Metadata = JsonUtils.safeStringify(alert.metadata);
        if ((alert.targetUsers?.length || 0) > 0) fields.TargetUsers = alert.targetUsers;
        
        fields.ItemType = alert.contentType;
        fields.TargetLanguage = alert.targetLanguage;
        if (alert.languageGroup) fields.LanguageGroup = alert.languageGroup;
        fields.AvailableForAll = Boolean(alert.availableForAll);

        let response = await this.graphClient.api(`${alertsListApi}/items`).post({ fields });
        
        // Retrieve created
        const created = await this.graphClient.api(`${alertsListApi}/items/${response.id}`).expand("fields").get();
        return AlertTransformers.mapSharePointItemToAlert(created, siteId);
      } catch (e) {
          logger.error("AlertOperationsService", "Create failed", e);
          throw e;
      }
  }

  public async updateAlert(alertId: string, updates: Partial<IAlertItem>): Promise<IAlertItem> {
      try {
          const { siteId, itemId } = this.parseAlertId(alertId);
          const alertsListApi = await this.locator.getAlertsListApi(siteId);

          const fields: any = {};
          if (updates.title) fields.Title = updates.title;
          if (updates.description) fields.Description = updates.description;
          if (updates.AlertType) fields.AlertType = updates.AlertType;
          if (updates.priority) fields.Priority = updates.priority;
          if (updates.isPinned !== undefined) fields.IsPinned = updates.isPinned;
          if (updates.notificationType) fields.NotificationType = updates.notificationType;
          if (updates.linkUrl !== undefined) fields.LinkUrl = updates.linkUrl;
          if (updates.linkDescription !== undefined) fields.LinkDescription = updates.linkDescription;
          if (updates.targetSites) fields.TargetSites = JsonUtils.safeStringify(updates.targetSites);
          if (updates.scheduledStart !== undefined) fields.ScheduledStart = updates.scheduledStart;
          if (updates.scheduledEnd !== undefined) fields.ScheduledEnd = updates.scheduledEnd;
          if (updates.targetUsers !== undefined) fields.TargetUsers = updates.targetUsers;
          if (updates.metadata) fields.Metadata = JsonUtils.safeStringify(updates.metadata);
          if (updates.status) fields.Status = updates.status;

          await this.graphClient.api(`${alertsListApi}/items/${itemId}/fields`).patch(fields);
          
          const updated = await this.graphClient.api(`${alertsListApi}/items/${itemId}`).expand("fields").get();
          return AlertTransformers.mapSharePointItemToAlert(updated, siteId);
      } catch (e) {
          logger.error("AlertOperationsService", "Update failed", e);
          throw e;
      }
  }

  public async deleteAlert(alertId: string): Promise<void> {
      try {
          const { siteId, itemId } = this.parseAlertId(alertId);
          const alertsListApi = await this.locator.getAlertsListApi(siteId);
          await this.graphClient.api(`${alertsListApi}/items/${itemId}`).delete();
      } catch (e) {
          logger.error("AlertOperationsService", "Delete failed", e);
          throw e;
      }
  }

  public async deleteAlerts(alertIds: string[]): Promise<void> {
      await Promise.allSettled(alertIds.map(id => this.deleteAlert(id)));
  }

  public parseAlertId(alertId: string): { siteId: string; itemId: string } {
      const lastHyphen = alertId.lastIndexOf("-");
      if (lastHyphen > 0 && lastHyphen < alertId.length - 1) {
          const site = alertId.substring(0, lastHyphen);
          const item = alertId.substring(lastHyphen + 1);
          if (/^\d+$/.test(item)) return { siteId: site, itemId: item };
      }
      return { siteId: this.context.pageContext.site.id.toString(), itemId: alertId };
  }

  public async getAlertTypes(siteIdOverride?: string): Promise<IAlertType[]> {
      const siteId = siteIdOverride || this.context.pageContext.site.id.toString();
      try {
          const listApi = await this.locator.getAlertTypesListApi(siteId); // Should Locator ensure existence? No. Provisioning does.
          // But here we just want to READ. If it fails, return defaults.
          const response = await this.graphClient.api(`${listApi}/items`).expand("fields").orderby("fields/SortOrder").get();
          
          if (!response.value || response.value.length === 0) return DEFAULT_ALERT_TYPES.map(t => ({ ...t }));
          
          return response.value.map((item: any) => ({
              name: item.fields.Title || "",
              iconName: item.fields.IconName || "Info",
              backgroundColor: item.fields.BackgroundColor || "#0078d4",
              textColor: item.fields.TextColor || "#ffffff",
              additionalStyles: item.fields.AdditionalStyles || "",
              priorityStyles: JsonUtils.safeParse(item.fields.PriorityStyles) || {}
          }));
      } catch (e) {
           return DEFAULT_ALERT_TYPES.map(t => ({ ...t }));
      }
  }


  public async getDraftAlerts(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.locator.getAlertsListApi(siteId);
      const currentUser = this.context.pageContext.user.loginName;
      const availableColumns = await this.locator.getAvailableColumns(alertsListApi);

      const baseFields = ["Title", "AlertType", "Description", "Priority", "IsPinned", "NotificationType", "LinkUrl", "LinkDescription", "TargetSites", "Status", "ItemType", "TargetLanguage", "LanguageGroup", "ScheduledStart", "ScheduledEnd", "TargetUsers", "Author", "Modified"];
      const selectedFields = baseFields.filter(f => availableColumns.has(f) || ["Title"].includes(f));
      
      const filters: string[] = [];
      if (availableColumns.has("ItemType")) filters.push("fields/ItemType eq 'draft'");
      if (availableColumns.has("Author") && currentUser) filters.push(`fields/Author/Email eq '${currentUser}'`);

      let request = this.graphClient.api(`${alertsListApi}/items`)
         .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
         .expand(`fields($select=${selectedFields.join(",")})`)
         .top(50);
      
      if (filters.length > 0) request = request.filter(filters.join(" and "));
      request = request.orderby(availableColumns.has("Modified") ? "fields/Modified desc" : "fields/Title asc");

      const response = await request.get();
      return response.value.map((item: any) => AlertTransformers.mapSharePointItemToAlert(item, siteId));
    } catch (e) {
      logger.warn("AlertOperationsService", "Failed to get drafts", e);
      return [];
    }
  }

  public async saveDraft(draft: Partial<IAlertItem>): Promise<IAlertItem> {
      const data: any = { ...draft, contentType: ContentType.Draft, status: "Draft" };
      if (draft.id && parseInt(draft.id) > 0) return this.updateAlert(draft.id, data);
      return this.createAlert(data);
  }

  public async deleteDraft(draftId: string): Promise<void> {
      return this.deleteAlert(draftId);
  }

  public async addAttachment(listId: string, itemId: number, fileName: string, fileContent: ArrayBuffer): Promise<{ fileName: string; serverRelativeUrl: string }> {
      try {
          const siteUrl = this.context.pageContext.web.absoluteUrl;
          const uploadUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`;
          const response = await this.context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
              headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/octet-stream' },
              body: fileContent
          });
          if (!response.ok) throw new Error(`Upload failed: ${response.statusText}`);
          const result = await response.json();
          return { fileName: result.d.FileName, serverRelativeUrl: result.d.ServerRelativeUrl };
      } catch (e) {
          logger.error("AlertOperationsService", "Upload attachment failed", e);
          throw e;
      }
  }

  public async deleteAttachment(listId: string, itemId: number, fileName: string): Promise<void> {
      try {
          const siteUrl = this.context.pageContext.web.absoluteUrl;
          const deleteUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
          await this.context.spHttpClient.post(deleteUrl, SPHttpClient.configurations.v1, {
              headers: { 'Accept': 'application/json;odata=verbose', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
          });
      } catch (e) {
          logger.error("AlertOperationsService", "Delete attachment failed", e);
          throw e;
      }
  }

  public async getAlertsForSite(siteId: string): Promise<IAlertItem[]> {
      try {
          const alertsListApi = await this.locator.getAlertsListApi(siteId);
          // Fetch all non-draft/template items
          const response = await this.graphClient.api(`${alertsListApi}/items`)
              .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
              .expand("fields")
              .top(50) 
              .orderby("fields/Created desc")
              .get();
          
          return response.value.filter((item: any) => {
              const type = (item.fields?.ItemType || "").toLowerCase();
              const status = (item.fields?.Status || "").toLowerCase();
              const title = item.fields?.Title || "";
              return type !== "draft" && type !== "template" && status !== "draft" && !title.startsWith("[Auto-saved]") && !title.startsWith("[auto-saved]");
          }).map((item: any) => AlertTransformers.mapSharePointItemToAlert(item, siteId));
      } catch (e) {
          logger.warn("AlertOperationsService", `Failed to get alerts for site ${siteId}`, e);
          return [];
      }
  }

  public async getTemplateAlerts(siteId: string): Promise<IAlertItem[]> {
      try {
          const alertsListApi = await this.locator.getAlertsListApi(siteId);
          // Simplified query
          const response = await this.graphClient.api(`${alertsListApi}/items`)
             .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
             .filter("fields/ItemType eq 'template'")
             .expand("fields")
             .get();
          return response.value.map((item: any) => AlertTransformers.mapSharePointItemToAlert(item, siteId));
      } catch (e) {
          logger.warn("AlertOperationsService", `Failed to get templates ${siteId}`, e);
          return [];
      }
  }

  public async updateAlertStatuses(siteIds?: string[]): Promise<void> {
    // Left empty for Facade to handle hierarchy-based updates
  }

  public async saveAlertTypes(alertTypes: IAlertType[]): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const alertTypesListApi = await this.locator.getAlertTypesListApi(siteId);

      // Clear existing items
      const existingItems = await this.graphClient.api(`${alertTypesListApi}/items`).expand("fields").get();
      for (const item of existingItems.value) {
          await this.graphClient.api(`${alertTypesListApi}/items/${item.id}`).delete();
      }

      // Add new items
      for (let i = 0; i < alertTypes.length; i++) {
        const alertType = alertTypes[i];
        await this.graphClient.api(`${alertTypesListApi}/items`).post({
          fields: {
             Title: alertType.name,
             IconName: alertType.iconName,
             BackgroundColor: alertType.backgroundColor,
             TextColor: alertType.textColor,
             AdditionalStyles: alertType.additionalStyles || "",
             PriorityStyles: JsonUtils.safeStringify(alertType.priorityStyles || {}) || "{}",
             SortOrder: i
          }
        });
      }
    } catch (error) {
       logger.error("AlertOperationsService", "Failed to save alert types", error);
       throw error;
    }
  }
}
