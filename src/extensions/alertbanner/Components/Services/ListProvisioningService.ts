import { MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SharePointListLocator } from "./SharePointListLocator";
import { logger } from "./LoggerService";
import { ErrorUtils } from "../Utils/ErrorUtils";
import { JsonUtils } from "../Utils/JsonUtils";
import { DateUtils } from "../Utils/DateUtils";
import { RetryUtils } from "../Utils/RetryUtils";
import {
  LIST_NAMES,
  SUPPORTED_LANGUAGES,
  DEFAULT_ALERT_TYPES,
  ALERT_ITEM_TYPES,
} from "../Utils/AppConstants";
import { AlertPriority, IAlertType, ContentType } from "../Alerts/IAlerts";

export interface IRepairResult {
  success: boolean;
  message: string;
  details: {
    columnsRemoved: string[];
    columnsAdded: string[];
    columnsUpdated: string[];
    errors: string[];
    warnings: string[];
  };
}

export class ListProvisioningService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private locator: SharePointListLocator;
  private alertsListName = LIST_NAMES.ALERTS;
  private alertTypesListName = LIST_NAMES.ALERT_TYPES;

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
    locator: SharePointListLocator,
  ) {
    this.graphClient = graphClient;
    this.context = context;
    this.locator = locator;
  }

  public async checkListsNeeded(): Promise<
    {
      site: string;
      needsAlerts: boolean;
      needsTypes: boolean;
      isHomeSite: boolean;
    }[]
  > {
    const results = [];
    const currentSiteId = this.context.pageContext.site.id.toString();
    const isHomeSite = await this.locator.isCurrentSite(currentSiteId);

    let isHome = false;
    try {
      const homeSiteResponse = await this.graphClient
        .api("/sites/root")
        .select("id")
        .get();
      const homeSiteId: string = homeSiteResponse.id;
      isHome =
        currentSiteId === homeSiteId ||
        (await this.locator.getGraphSiteIdentifierFromContext(
          currentSiteId,
        )) ===
          (await this.locator.getGraphSiteIdentifierFromContext(homeSiteId));
    } catch (error) {
      logger.warn(
        "ListProvisioningService",
        "Unable to check home site status",
        error,
      );
    }

    // Check current site
    let needsAlerts = false;
    let needsTypes = false;

    try {
      await this.locator.resolveListId(currentSiteId, this.alertsListName);
    } catch (error: unknown) {
      if (ErrorUtils.isListNotFoundError(error)) {
        needsAlerts = true;
      } else if (!ErrorUtils.isAccessDeniedError(error)) {
        throw error;
      }
    }

    if (isHome) {
      try {
        await this.locator.resolveListId(
          currentSiteId,
          this.alertTypesListName,
        );
      } catch (error: unknown) {
        if (ErrorUtils.isListNotFoundError(error)) {
          needsTypes = true;
        } else if (!ErrorUtils.isAccessDeniedError(error)) {
          throw error;
        }
      }
    }

    results.push({
      site: currentSiteId,
      needsAlerts,
      needsTypes,
      isHomeSite: isHome,
    });

    return results;
  }

  public async initializeLists(siteId?: string): Promise<void> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();

    // Re-implement isHomeSite logic or optimize
    let isHomeSite = false;
    try {
      const homeSiteResponse = await this.graphClient
        .api("/sites/root")
        .select("id")
        .get();
      // Compare with normalized IDs
      const normalizedTarget = (
        await this.locator.ensureGraphSiteIdentifier(targetSiteId)
      ).toLowerCase();
      const normalizedHome = homeSiteResponse.id.toLowerCase();

      isHomeSite =
        targetSiteId === homeSiteResponse.id ||
        normalizedTarget === normalizedHome;
    } catch (e) {
      logger.warn("ListProvisioningService", "Home site check failed", e);
    }

    try {
      await this.ensureAlertsList(targetSiteId);
    } catch (alertsError: unknown) {
      if (
        alertsError instanceof Error &&
        alertsError.message?.includes("PERMISSION_DENIED")
      ) {
        logger.warn(
          "ListProvisioningService",
          "Cannot create alerts list due to insufficient permissions",
        );
      } else {
        throw alertsError;
      }
    }

    if (isHomeSite) {
      try {
        await this.ensureAlertTypesList(targetSiteId);
      } catch (typesError: unknown) {
        if (
          typesError instanceof Error &&
          typesError.message?.includes("PERMISSION_DENIED")
        ) {
          logger.warn(
            "ListProvisioningService",
            "Cannot create types list on home site due to insufficient permissions",
          );
        } else {
          throw typesError;
        }
      }
    }
  }

  private async ensureAlertsList(siteId: string): Promise<boolean> {
    try {
      await this.locator.resolveListId(siteId, this.alertsListName);
      return false;
    } catch (error: unknown) {
      if (ErrorUtils.isAccessDeniedError(error)) {
        logger.warn(
          "ListProvisioningService",
          "Access denied checking alerts list",
        );
        throw new Error(
          "PERMISSION_DENIED: User lacks permissions to access or create SharePoint lists.",
        );
      }
      if (!ErrorUtils.isListNotFoundError(error)) {
        throw error;
      }
    }

    const graphSiteIdentifier =
      await this.locator.ensureGraphSiteIdentifier(siteId);

    // Permission check
    try {
      await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .select("id")
        .top(1)
        .get();
    } catch (permissionError: unknown) {
      if (ErrorUtils.isAccessDeniedError(permissionError)) {
        throw new Error(
          "PERMISSION_DENIED: User lacks permissions to create SharePoint lists.",
        );
      }
    }

    const listDefinition = {
      displayName: this.alertsListName,
      list: { template: "genericList", contentTypesEnabled: false },
    };

    try {
      const createdList = await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .post(listDefinition);

      await this.locator.registerListId(
        siteId,
        this.alertsListName,
        createdList?.id,
      );
      await this.enableListAttachments(siteId, createdList?.id);
      await this.addAlertsListColumns(siteId);
      await this.addAlertsListColumns(siteId);
      await this.seedDefaultAlertTypes(siteId);

      await this.createTemplateAlerts(siteId);

      return true;
    } catch (createError: unknown) {
      if (ErrorUtils.isAccessDeniedError(createError))
        throw new Error("PERMISSION_DENIED");
      throw createError;
    }
  }

  private async ensureAlertTypesList(siteId: string): Promise<boolean> {
    try {
      await this.locator.resolveListId(siteId, this.alertTypesListName);
      return false;
    } catch (error: unknown) {
      if (!ErrorUtils.isListNotFoundError(error)) throw error;
    }

    const graphSiteIdentifier =
      await this.locator.ensureGraphSiteIdentifier(siteId);

    const listDefinition = {
      displayName: this.alertTypesListName,
      list: { template: "genericList" },
    };

    try {
      const createdList = await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .post(listDefinition);
      await this.locator.registerListId(
        siteId,
        this.alertTypesListName,
        createdList?.id,
      );
      await this.addAlertTypesListColumns(siteId);
      await this.seedDefaultAlertTypes(siteId);
      return true;
    } catch (error) {
      throw error;
    }
  }

  private async enableListAttachments(
    siteId: string,
    listId: string,
  ): Promise<void> {
    try {
      const siteUrl = await this.locator.getSiteUrlFromIdentifier(siteId);
      const updateUrl = `${siteUrl}/_api/web/lists(guid'${listId}')`;

      await this.context.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({
            __metadata: { type: "SP.List" },
            EnableAttachments: true,
          }),
        },
      );
      logger.info("ListProvisioningService", "Attachments enabled");
    } catch (error) {
      logger.warn(
        "ListProvisioningService",
        "Failed to enable attachments",
        error,
      );
    }
  }

  private async addAlertsListColumns(
    siteId: string,
    onlyMissing?: Set<string>,
  ): Promise<void> {
    let alertTypesListId = "";
    try {
      alertTypesListId = await this.locator.resolveListId(
        siteId,
        this.alertTypesListName,
      );
    } catch (e) {
      /* Ignore */
    }

    const alertsListId = await this.locator.resolveListId(
      siteId,
      this.alertsListName,
    );
    const graphSiteIdentifier =
      await this.locator.ensureGraphSiteIdentifier(siteId);

    const columns = [
      alertTypesListId
        ? {
            name: "AlertType",
            lookup: {
              listId: alertTypesListId,
              columnName: "Title",
              allowMultipleValues: false,
              allowUnlimitedLength: false,
            },
          }
        : {
            name: "AlertType",
            text: { maxLength: 255 },
          },
      {
        name: "Description",
        text: { allowMultipleLines: true, maxLength: 4000 },
      },
      {
        name: "Priority",
        choice: {
          choices: ["low", "medium", "high", "critical"],
          displayAs: "dropdown",
        },
      },
      { name: "IsPinned", boolean: {} },
      {
        name: "NotificationType",
        choice: {
          choices: ["none", "browser", "email", "both"],
          displayAs: "dropdown",
        },
      },
      { name: "LinkUrl", text: {} },
      { name: "LinkDescription", text: { maxLength: 255 } },
      {
        name: "TargetSites",
        text: { allowMultipleLines: true, maxLength: 4000 },
      },
      {
        name: "Status",
        choice: {
          choices: ["Active", "Expired", "Scheduled", "Draft"],
          displayAs: "dropdown",
        },
      },
      {
        name: "ScheduledStart",
        dateTime: { displayAs: "default", format: "dateTime" },
      },
      {
        name: "ScheduledEnd",
        dateTime: { displayAs: "default", format: "dateTime" },
      },
      { name: "Metadata", text: { allowMultipleLines: true, maxLength: 4000 } },
      {
        name: "ItemType",
        choice: {
          choices: [
            ALERT_ITEM_TYPES.ALERT,
            ALERT_ITEM_TYPES.TEMPLATE,
            ALERT_ITEM_TYPES.DRAFT,
            ALERT_ITEM_TYPES.SETTINGS,
          ],
          displayAs: "dropdown",
        },
        indexed: true,
      },
      {
        name: "TargetLanguage",
        choice: {
          choices: ["all", ...SUPPORTED_LANGUAGES.map((l) => l.code)],
          displayAs: "dropdown",
        },
      },
      { name: "LanguageGroup", text: { maxLength: 255 } },
      { name: "AvailableForAll", boolean: {} },
      {
        name: "TranslationStatus",
        choice: {
          choices: ["Draft", "InReview", "Approved"],
          displayAs: "dropdown",
        },
      },
      {
        name: "TargetUsers",
        personOrGroup: {
          allowMultipleSelection: true,
          chooseFromType: "peopleAndGroups",
        },
      },
      {
        name: "ContentStatus",
        choice: {
          choices: ["Draft", "PendingReview", "Approved", "Rejected"],
          displayAs: "dropdown",
        },
      },
      {
        name: "Reviewer",
        personOrGroup: {
          allowMultipleSelection: true,
          chooseFromType: "peopleAndGroups",
        },
      },
      {
        name: "ReviewNotes",
        text: { allowMultipleLines: true, maxLength: 4000 },
      },
      {
        name: "SubmittedDate",
        dateTime: { displayAs: "default", format: "dateTime" },
      },
      {
        name: "ReviewedDate",
        dateTime: { displayAs: "default", format: "dateTime" },
      },
      {
        name: "SortOrder",
        number: { decimalPlaces: "none" },
        indexed: true,
      },
    ];

    for (const column of columns) {
      if (onlyMissing && !onlyMissing.has(column.name)) {
        continue;
      }
      try {
        await this.graphClient
          .api(`/sites/${graphSiteIdentifier}/lists/${alertsListId}/columns`)
          .post(column);
      } catch (error) {
        logger.warn(
          "ListProvisioningService",
          `Failed to create column ${column.name}`,
          error,
        );
      }
    }
  }

  private async addAlertTypesListColumns(siteId: string): Promise<void> {
    const columns = [
      { name: "IconName", text: { maxLength: 100 } },
      { name: "BackgroundColor", text: { maxLength: 50 } },
      { name: "TextColor", text: { maxLength: 50 } },
      {
        name: "AdditionalStyles",
        text: { allowMultipleLines: true, maxLength: 4000 },
      },
      {
        name: "PriorityStyles",
        text: { allowMultipleLines: true, maxLength: 4000 },
      },
      { name: "SortOrder", number: { decimalPlaces: "none" }, indexed: true },
    ];

    const listApi = await this.locator.getAlertTypesListApi(siteId);
    for (const column of columns) {
      try {
        await this.graphClient.api(`${listApi}/columns`).post(column);
      } catch (e) {
        logger.warn(
          "ListProvisioningService",
          `Failed column ${column.name}`,
          e,
        );
      }
    }
  }

  private async seedDefaultAlertTypes(siteId: string): Promise<void> {
    try {
      const listApi = await this.locator.getAlertTypesListApi(siteId);
      const existing = await this.graphClient
        .api(`${listApi}/items`)
        .top(1)
        .get();
      if (existing.value?.length > 0) return;

      const defaults = DEFAULT_ALERT_TYPES;
      let sortOrder = 0;
      for (const type of defaults) {
        // Include defaultPriority in PriorityStyles JSON
        const priorityStylesPayload = {
          ...(type.priorityStyles || {}),
          __defaultPriority: type.defaultPriority || undefined,
        };
        const payload = {
          fields: {
            Title: type.name,
            IconName: type.iconName,
            BackgroundColor: type.backgroundColor,
            TextColor: type.textColor,
            AdditionalStyles: type.additionalStyles || "",
            PriorityStyles:
              JsonUtils.safeStringify(priorityStylesPayload) || "{}",
            SortOrder: sortOrder++,
          },
        };
        try {
          await this.graphClient.api(`${listApi}/items`).post(payload);
        } catch (e) {
          logger.warn(
            "ListProvisioningService",
            `Failed to seed ${type.name}`,
            e,
          );
        }
      }
    } catch (e) {
      logger.warn("ListProvisioningService", "Seed failed", e);
    }
  }

  private async createTemplateAlerts(siteId: string): Promise<void> {
    try {
      const defaultTemplates = require("../Data/defaultTemplates.json");
      const alertsListApi = await this.locator.getAlertsListApi(siteId);

      const templateAlerts = defaultTemplates.map(
        (template: {
          fields: Record<string, unknown>;
          [key: string]: unknown;
        }) => ({
          ...template,
          fields: {
            ...template.fields,
            ScheduledStart: new Date().toISOString(),
            ScheduledEnd: this.getTemplateEndDate(
              String(template.fields.AlertType || ""),
            ),
            ItemType: template.fields.ContentType,
            ContentType: undefined,
          },
        }),
      );

      for (const template of templateAlerts) {
        try {
          await this.graphClient.api(`${alertsListApi}/items`).post(template);
        } catch (e) {
          logger.warn("ListProvisioningService", "Template creation failed", e);
        }
      }
    } catch (e) {
      logger.warn("ListProvisioningService", "Failed to load templates", e);
    }
  }

  private getTemplateEndDate(alertType: string): string {
    const now = new Date();
    switch (alertType.toLowerCase()) {
      case "maintenance":
        return DateUtils.addDurationISO(now, 1, "days");
      case "warning":
        return DateUtils.addDurationISO(now, 3, "days");
      case "interruption":
        return DateUtils.addDurationISO(now, 12, "hours");
      case "info":
        return DateUtils.addDurationISO(now, 1, "weeks");
      default:
        return DateUtils.addDurationISO(now, 1, "months");
    }
  }

  public async repairAlertsList(
    siteId: string,
    progressCallback?: (msg: string, p: number) => void,
  ): Promise<IRepairResult> {
    const result: IRepairResult = {
      success: false,
      message: "",
      details: {
        columnsRemoved: [],
        columnsAdded: [],
        columnsUpdated: [],
        errors: [],
        warnings: [],
      },
    };

    try {
      progressCallback?.("Analyzing list schema...", 10);

      const alertsListApi = await this.locator.getAlertsListApi(siteId);
      if (!alertsListApi) {
        throw new Error(
          `Alerts list not found for site ${siteId}. Please provision the list first.`,
        );
      }

      const alertsListId = await this.locator.resolveListId(
        siteId,
        this.alertsListName,
      );

      const currentColumnsResponse = await this.graphClient
        .api(`${alertsListApi}/columns`)
        .get();
      const currentColumns = currentColumnsResponse?.value || [];

      const expectedColumns = [
        "Title",
        "Description",
        "AlertType",
        "Priority",
        "IsPinned",
        "NotificationType",
        "LinkUrl",
        "LinkDescription",
        "TargetSites",
        "Status",
        "ScheduledStart",
        "ScheduledEnd",
        "Metadata",
        "ItemType",
        "TargetLanguage",
        "LanguageGroup",
        "AvailableForAll",
        "TranslationStatus",
        "TargetUsers",
        "ContentStatus",
        "Reviewer",
        "ReviewNotes",
        "SubmittedDate",
        "ReviewedDate",
      ];

      const coreColumns = new Set([
        "Title",
        "Created",
        "Modified",
        "Author",
        "Editor",
        "ID",
        "Attachments",
        "Body",
      ]);

      const expectedSet = new Set(expectedColumns);

      progressCallback?.("Removing obsolete columns...", 30);

      const customColumns = currentColumns.filter(
        (col: { readOnly?: boolean; name?: string; id?: string }) =>
          !col.readOnly &&
          col.name &&
          !col.name.startsWith("_") &&
          !coreColumns.has(col.name),
      );

      for (const column of customColumns) {
        if (!expectedSet.has(column.name)) {
          try {
            await this.graphClient
              .api(`${alertsListApi}/columns/${column.id}`)
              .delete();
            result.details.columnsRemoved.push(column.name);
          } catch (e: unknown) {
            const msg = e instanceof Error ? e.message : String(e);
            result.details.warnings.push(
              `Failed to remove ${column.name}: ${msg}`,
            );
          }
        }
      }

      progressCallback?.("Adding missing columns...", 55);

      const existingNames = new Set<string>(
        currentColumns.map((col: { name: string }) => col.name),
      );
      const missingColumns = expectedColumns.filter(
        (name) => !existingNames.has(name),
      );
      const missingSet = new Set(missingColumns);

      if (missingColumns.length > 0) {
        await this.addAlertsListColumns(siteId, missingSet);
        result.details.columnsAdded.push(...missingColumns);
      }

      progressCallback?.("Updating choice values...", 70);

      const getColumnByName = (name: string) =>
        currentColumns.find((col: { name: string }) => col.name === name);

      const ensureChoices = async (
        columnName: string,
        expectedChoices: string[],
      ) => {
        const column = getColumnByName(columnName);
        if (!column || (!column.choice && !column.choices)) {
          return;
        }

        const currentChoices = column.choice?.choices || column.choices || [];
        const missing = expectedChoices.filter(
          (choice) => !currentChoices.includes(choice),
        );

        if (missing.length === 0) {
          return;
        }

        const updatedChoices = Array.from(
          new Set([...currentChoices, ...missing]),
        );
        try {
          await this.graphClient
            .api(`${alertsListApi}/columns/${column.id}`)
            .patch({
              choice: { ...(column.choice || {}), choices: updatedChoices },
            });
          result.details.columnsUpdated.push(
            `${columnName}: added ${missing.join(", ")}`,
          );
        } catch (e: unknown) {
          const msg = e instanceof Error ? e.message : String(e);
          result.details.warnings.push(
            `Failed to update ${columnName} choices: ${msg}`,
          );
        }
      };

      await ensureChoices("Priority", ["low", "medium", "high", "critical"]);
      await ensureChoices("NotificationType", [
        "none",
        "browser",
        "email",
        "both",
      ]);
      await ensureChoices("Status", [
        "Active",
        "Expired",
        "Scheduled",
        "Draft",
      ]);
      await ensureChoices("ItemType", [
        ALERT_ITEM_TYPES.ALERT,
        ALERT_ITEM_TYPES.TEMPLATE,
        ALERT_ITEM_TYPES.DRAFT,
        ALERT_ITEM_TYPES.SETTINGS,
      ]);
      await ensureChoices("TargetLanguage", [
        "all",
        ...SUPPORTED_LANGUAGES.map((l) => l.code),
      ]);
      await ensureChoices("TranslationStatus", [
        "Draft",
        "InReview",
        "Approved",
      ]);
      await ensureChoices("ContentStatus", [
        "Draft",
        "PendingReview",
        "Approved",
        "Rejected",
      ]);

      progressCallback?.("Ensuring indexes and settings...", 85);

      await this.ensureItemTypeIndex(siteId, alertsListApi);
      await this.enableListAttachments(siteId, alertsListId);

      result.success = true;
      result.message = "Alerts list repaired successfully.";
      progressCallback?.("Repair completed.", 100);
      return result;
    } catch (e: unknown) {
      result.message = e instanceof Error ? e.message : "Repair failed";
      result.details.errors.push(result.message);
      return result;
    }
  }

  private async ensureItemTypeIndex(
    siteId: string,
    alertsListApi: string,
  ): Promise<void> {
    try {
      const columns = await this.graphClient
        .api(`${alertsListApi}/columns`)
        .select("id,name,indexed")
        .get();
      const itemType = columns.value?.find(
        (c: { name: string }) => c.name === "ItemType",
      );

      if (itemType && !itemType.indexed) {
        await this.graphClient
          .api(`${alertsListApi}/columns/${itemType.id}`)
          .patch({ indexed: true });
        logger.info("ListProvisioningService", "Indexed ItemType column");
      }
    } catch (e) {
      logger.warn("ListProvisioningService", "Failed to index ItemType", e);
    }
  }

  public async getSupportedLanguages(siteId?: string): Promise<string[]> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();
    try {
      const alertsListApi = await this.locator.getAlertsListApi(targetSiteId);
      const columnsResponse = await this.graphClient
        .api(`${alertsListApi}/columns`)
        .select("name,choice")
        .get();

      const targetLangColumn = (columnsResponse.value || []).find(
        (col: { name?: string }) =>
          (col.name || "").toLowerCase() === "targetlanguage",
      );
      const choices: string[] = targetLangColumn?.choice?.choices ||
        targetLangColumn?.choices || ["en-us"];
      return choices.filter((c) => c.toLowerCase() !== "all");
    } catch (e) {
      logger.warn("ListProvisioningService", "Failed to get languages", e);
      return ["en-us"];
    }
  }

  public async addLanguageSupport(
    languageCode: string,
    siteId?: string,
  ): Promise<void> {
    await this.updateTargetLanguageChoices("add", languageCode, siteId);
  }

  public async removeLanguageSupport(languageCode: string): Promise<void> {
    await this.updateTargetLanguageChoices("remove", languageCode);
  }

  public async updateSupportedLanguages(
    siteId: string,
    enabledLanguages: string[],
  ): Promise<void> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();
    const current = await this.getSupportedLanguages(targetSiteId);
    const toAdd = enabledLanguages.filter(
      (l) => !current.includes(l.toLowerCase()),
    );
    for (const lang of toAdd) {
      if (lang !== "en-us") await this.addLanguageSupport(lang, targetSiteId);
    }
  }

  private async updateTargetLanguageChoices(
    action: "add" | "remove",
    languageCode: string,
    siteId?: string,
  ): Promise<void> {
    const targetSiteId = siteId || this.context.pageContext.site.id.toString();
    const alertsListApi = await this.locator.getAlertsListApi(targetSiteId);
    const columns = await this.graphClient
      .api(`${alertsListApi}/columns`)
      .select("id,name,choice")
      .get();

    let column = (columns.value || []).find(
      (c: any) => (c.name || "").toLowerCase() === "targetlanguage",
    );

    // Create if missing (Provisioning logic!)
    if (!column) {
      await this.graphClient.api(`${alertsListApi}/columns`).post({
        name: "TargetLanguage",
        choice: {
          allowTextEntry: false,
          choices: ["all", "en-us"],
          displayAs: "dropdown",
        },
      });
      const refreshed = await this.graphClient
        .api(`${alertsListApi}/columns`)
        .get();
      column = (refreshed.value || []).find(
        (c: { name?: string }) =>
          (c.name || "").toLowerCase() === "targetlanguage",
      );
    }

    const currentChoices = column.choice?.choices ||
      column.choices || ["all", "en-us"];
    let updated: string[];

    if (action === "add") {
      if (currentChoices.includes(languageCode)) return;
      updated = [...currentChoices, languageCode].sort();
    } else {
      updated = currentChoices.filter(
        (c: string) => c !== languageCode || c === "all" || c === "en-us",
      );
      if (updated.length === currentChoices.length) return;
    }

    await this.graphClient.api(`${alertsListApi}/columns/${column.id}`).patch({
      choice: { ...column.choice, choices: updated },
    });
  }
}
