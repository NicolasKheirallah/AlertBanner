import {
  AlertPriority,
  ContentStatus,
  ContentType,
  IAlertItem,
  IPersonField,
  NotificationType,
  TargetLanguage,
  TranslationStatus,
  ILanguageContent,
} from "../Alerts/IAlerts";
import { LanguageAwarenessService } from "./LanguageAwarenessService";
import { ILanguagePolicy } from "./LanguagePolicyService";
import { SharePointAlertService } from "./SharePointAlertService";
import { logger } from "./LoggerService";

export interface IAlertMutationModel {
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string[];
  scheduledStart?: Date;
  scheduledEnd?: Date;
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageGroup?: string;
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
  languageContent?: ILanguageContent[];
}

export interface IEditableAlertMutationModel extends IAlertMutationModel {
  id: string;
}

interface ICreateAlertsOptions {
  useMultiLanguage: boolean;
  createdBy: string;
  languagePolicy: ILanguagePolicy;
  languageService: LanguageAwarenessService;
}

interface IDraftPayloadOptions {
  id?: string;
  createdBy: string;
  titlePrefix?: string;
  contentStatus?: ContentStatus;
}

export class AlertMutationService {
  public static async createAlertsFromModel(
    alertService: SharePointAlertService,
    alert: IAlertMutationModel,
    options: ICreateAlertsOptions,
  ): Promise<number> {
    if (
      options.useMultiLanguage &&
      alert.languageContent &&
      alert.languageContent.length > 0
    ) {
      const multiLanguageAlert =
        options.languageService.createMultiLanguageAlert(
          {
            AlertType: alert.AlertType,
            priority: alert.priority,
            isPinned: alert.isPinned,
            linkUrl: alert.linkUrl || "",
            notificationType: alert.notificationType,
            createdDate: new Date().toISOString(),
            createdBy: options.createdBy,
            contentType: alert.contentType,
            contentStatus: this.getContentStatus(alert.contentType),
            targetLanguage: TargetLanguage.All,
            status: "Active" as IAlertItem["status"],
            targetSites: alert.targetSites || [],
            id: "0",
            targetUsers: this.getCombinedTargets(alert),
            languageGroup: alert.languageGroup,
          },
          alert.languageContent,
        );

      const alertItems =
        options.languageService.generateAlertItems(multiLanguageAlert);
      const createdVariantIds: string[] = [];

      try {
        for (const alertItem of alertItems) {
          const createdAlert = await alertService.createAlert({
            ...alertItem,
            targetSites: alert.targetSites || [],
            scheduledStart: alert.scheduledStart?.toISOString(),
            scheduledEnd: alert.scheduledEnd?.toISOString(),
          });
          createdVariantIds.push(createdAlert.id);
        }
      } catch (error) {
        const rollbackFailures = await this.rollbackCreatedAlerts(
          alertService,
          createdVariantIds,
        );
        const rollbackDetails =
          rollbackFailures.length > 0
            ? ` Rollback failures: ${rollbackFailures.join("; ")}`
            : "";
        throw new Error(
          `Failed to create multi-language alert variants: ${this.getErrorMessage(error)}.${rollbackDetails}`,
        );
      }

      return alertItems.length;
    }

    await alertService.createAlert({
      ...this.toBaseMutationFields(alert),
      contentType: alert.contentType,
      contentStatus: this.getContentStatus(alert.contentType),
      targetLanguage: alert.targetLanguage,
      languageGroup: alert.languageGroup,
      translationStatus: this.getDefaultTranslationStatus(
        options.languagePolicy,
      ),
    });

    return 1;
  }

  public static async updateSingleAlert(
    alertService: SharePointAlertService,
    alert: IEditableAlertMutationModel,
  ): Promise<void> {
    const payload = this.toBaseMutationFields(alert);
    logger.debug("AlertMutationService", "Updating single alert", {
      alertId: alert.id,
      alertType: payload.AlertType,
      priority: payload.priority,
      title: payload.title,
    });
    await alertService.updateAlert(alert.id, payload);
  }

  public static async syncMultiLanguageAlerts(
    alertService: SharePointAlertService,
    editingAlert: IEditableAlertMutationModel,
    existingAlerts: IAlertItem[],
    languagePolicy: ILanguagePolicy,
  ): Promise<void> {
    if (!editingAlert.languageGroup) {
      logger.warn(
        "AlertMutationService",
        "Multi-language sync requested without languageGroup; falling back to single update",
        { id: editingAlert.id },
      );
      await this.updateSingleAlert(alertService, editingAlert);
      return;
    }

    const languageContent = editingAlert.languageContent || [];
    if (languageContent.length === 0) {
      await this.updateSingleAlert(alertService, editingAlert);
      return;
    }

    const allGroupAlerts = existingAlerts.filter(
      (a) => a.languageGroup === editingAlert.languageGroup,
    );

    const dedupLanguages = new Set<string>();
    const groupAlerts = allGroupAlerts.filter((alert) => {
      if (dedupLanguages.has(alert.targetLanguage)) {
        return false;
      }
      dedupLanguages.add(alert.targetLanguage);
      return true;
    });

    const defaultTranslationStatus =
      this.getDefaultTranslationStatus(languagePolicy);
    const createdVariantIds: string[] = [];

    for (const content of languageContent) {
      const existingAlert = groupAlerts.find(
        (a) => a.targetLanguage === content.language,
      );

      const variantPayload = {
        ...this.toBaseMutationFields({
          ...editingAlert,
          title: content.title,
          description: content.description,
          linkDescription: content.linkDescription || "",
        }),
        availableForAll: content.availableForAll,
        translationStatus:
          content.translationStatus || defaultTranslationStatus,
      };

      if (existingAlert) {
        try {
          await alertService.updateAlert(existingAlert.id, variantPayload);
        } catch (error) {
          throw new Error(
            `Failed to update language variant '${content.language}': ${this.getErrorMessage(error)}`,
          );
        }
      } else {
        try {
          const createdAlert = await alertService.createAlert({
            ...variantPayload,
            contentType: editingAlert.contentType,
            targetLanguage: content.language,
            languageGroup: editingAlert.languageGroup,
          });
          createdVariantIds.push(createdAlert.id);
        } catch (error) {
          const rollbackFailures = await this.rollbackCreatedAlerts(
            alertService,
            createdVariantIds,
          );
          const rollbackDetails =
            rollbackFailures.length > 0
              ? ` Rollback failures: ${rollbackFailures.join("; ")}`
              : "";
          throw new Error(
            `Failed to create language variant '${content.language}': ${this.getErrorMessage(error)}.${rollbackDetails}`,
          );
        }
      }
    }

    const updatedLanguages = languageContent.map((c) => c.language);
    const seenLanguages = new Set<string>();
    const toDelete = allGroupAlerts.filter((a) => {
      if (!updatedLanguages.includes(a.targetLanguage)) {
        return true;
      }
      if (seenLanguages.has(a.targetLanguage)) {
        return true;
      }
      seenLanguages.add(a.targetLanguage);
      return false;
    });

    const deletionFailures: string[] = [];
    for (const alertToDelete of toDelete) {
      try {
        await alertService.deleteAlert(alertToDelete.id);
      } catch (deleteError: unknown) {
        const err = deleteError as {
          statusCode?: number;
          status?: number;
          response?: { status?: number };
          code?: string;
          message?: string;
        };
        const statusCode =
          err?.statusCode || err?.status || err?.response?.status || err?.code;
        const errorMessage = err?.message || String(deleteError);

        const isNotFound =
          statusCode === 404 ||
          statusCode === "404" ||
          errorMessage.toLowerCase().includes("not found") ||
          errorMessage.toLowerCase().includes("does not exist");

        if (!isNotFound) {
          deletionFailures.push(
            `${alertToDelete.id} (${alertToDelete.targetLanguage}): ${this.getErrorMessage(deleteError)}`,
          );
        }
      }
    }

    if (deletionFailures.length > 0) {
      throw new Error(
        `Updated alert content but failed deleting removed language variants: ${deletionFailures.join("; ")}`,
      );
    }
  }

  public static buildDraftPayload(
    alert: IAlertMutationModel,
    options: IDraftPayloadOptions,
  ): Partial<IAlertItem> {
    return {
      id: options.id,
      ...this.toBaseMutationFields({
        ...alert,
        title: `${options.titlePrefix || ""}${alert.title}`,
      }),
      contentType: ContentType.Draft,
      contentStatus: options.contentStatus || ContentStatus.Draft,
      targetLanguage: alert.targetLanguage,
      languageGroup: alert.languageGroup,
      createdDate: new Date().toISOString(),
      createdBy: options.createdBy,
    };
  }

  private static getCombinedTargets(
    alert: Pick<IAlertMutationModel, "targetUsers" | "targetGroups">,
  ): IPersonField[] {
    return [...(alert.targetUsers || []), ...(alert.targetGroups || [])];
  }

  private static getContentStatus(contentType: ContentType): ContentStatus {
    return contentType === ContentType.Alert
      ? ContentStatus.Approved
      : ContentStatus.Draft;
  }

  private static getDefaultTranslationStatus(
    languagePolicy: ILanguagePolicy,
  ): TranslationStatus {
    return languagePolicy.workflow.enabled
      ? languagePolicy.workflow.defaultStatus
      : TranslationStatus.Approved;
  }

  private static toBaseMutationFields(
    alert: IAlertMutationModel,
  ): Partial<IAlertItem> {
    return {
      title: alert.title,
      description: alert.description,
      AlertType: alert.AlertType,
      priority: alert.priority,
      isPinned: alert.isPinned,
      notificationType: alert.notificationType,
      linkUrl: alert.linkUrl,
      linkDescription: alert.linkDescription,
      scheduledStart: alert.scheduledStart?.toISOString(),
      scheduledEnd: alert.scheduledEnd?.toISOString(),
      targetSites: alert.targetSites || [],
      targetUsers: this.getCombinedTargets(alert),
    };
  }

  private static getErrorMessage(error: unknown): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }
    return String(error);
  }

  private static async rollbackCreatedAlerts(
    alertService: SharePointAlertService,
    createdVariantIds: string[],
  ): Promise<string[]> {
    const failures: string[] = [];
    for (const id of createdVariantIds) {
      try {
        await alertService.deleteAlert(id);
      } catch (rollbackError) {
        failures.push(`${id}: ${this.getErrorMessage(rollbackError)}`);
      }
    }
    return failures;
  }
}
