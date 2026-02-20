import * as React from "react";
import {
  Save24Regular,
  Dismiss24Regular,
  Drafts24Regular,
  DocumentArrowLeft24Regular,
  ArrowLeft24Regular,
  ArrowRight24Regular,
  History24Regular,
  Delete24Regular,
} from "@fluentui/react-icons";
import {
  SharePointButton,
  ISharePointSelectOption,
} from "../../UI/SharePointControls";
import type { IAlertTemplate } from "../../UI/AlertTemplates";
import {
  AlertPriority,
  NotificationType,
  ContentType,
  TargetLanguage,
  IAlertItem,
  TranslationStatus,
} from "../../Alerts/IAlerts";
import { LanguageAwarenessService } from "../../Services/LanguageAwarenessService";
import { DEFAULT_LANGUAGE_POLICY } from "../../Services/LanguagePolicyService";
import { logger } from "../../Services/LoggerService";
import { NotificationService } from "../../Services/NotificationService";
import { CopilotService } from "../../Services/CopilotService";
import styles from "../AlertSettings.module.scss";
import { validateAlertData } from "../../Utils/AlertValidation";
import { getLocalizedValidationMessage } from "../../Utils/AlertValidationLocalization";
import { useLanguageOptions } from "../../Hooks/useLanguageOptions";
import { DateUtils } from "../../Utils/DateUtils";
import { StringUtils } from "../../Utils/StringUtils";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import AlertEditorForm, { CreateWizardStep } from "./AlertEditorForm";
import { AlertMutationService } from "../../Services/AlertMutationService";
import { useFluentDialogs } from "../../Hooks/useFluentDialogs";
import EmptyState from "../../UI/EmptyState";
import {
  useAlertForm,
  useAlertFormState,
  useAlertFormServices,
  INewAlert,
} from "../../Context/AlertFormContext";

const AlertTemplates = React.lazy(() => import("../../UI/AlertTemplates"));

export type { INewAlert };
export type { IFormErrors } from "./SharedTypes";

/**
 * Props for the CreateAlertTab component
 * Most state is now managed via AlertFormContext
 */
export interface ICreateAlertTabProps {
  /** Optional callback when dirty state changes */
  onDirtyStateChange?: (hasUnsavedChanges: boolean) => void;
}


/**
 * Inner component that uses the context
 * Separated to allow proper hook usage within provider
 */
const CreateAlertTabInner: React.FC<ICreateAlertTabProps> = ({
  onDirtyStateChange,
}) => {
  const { state, dispatch } = useAlertForm();
  const services = useAlertFormServices();

  const {
    newAlert,
    errors,
    isCreatingAlert,
    showPreview,
    showTemplates,
    createStep,
    currentEntryMode,
    creationProgress,
    lastCreateAttemptFailed,
    autoSaveDraftId,
    lastAutoSave,
    autoSaveStatus,
    isAutoSaving,
    useMultiLanguage,
    supportedLanguages,
    languagePolicy,
    drafts,
    previousAlerts,
    copilotAvailability,
    alertTypes,
    userTargetingEnabled,
    notificationsEnabled,
    enableTargetSite,
    copilotEnabled,
  } = state;

  const { siteDetector, alertService, graphClient, context, languageService } = services;

  const notificationOptions: ISharePointSelectOption[] = React.useMemo(
    () => [
      {
        value: NotificationType.None,
        label: strings.CreateAlertNotificationNoneDescription,
      },
      {
        value: NotificationType.Browser,
        label: strings.CreateAlertNotificationBrowserDescription,
      },
      {
        value: NotificationType.Email,
        label: strings.CreateAlertNotificationEmailDescription,
      },
      {
        value: NotificationType.Both,
        label: strings.CreateAlertNotificationBothDescription,
      },
    ],
    [],
  );

  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(
    (type) => ({
      value: type.name,
      label: type.name,
    }),
  );

  const contentTypeOptions: ISharePointSelectOption[] = React.useMemo(
    () => [
      {
        value: ContentType.Alert,
        label: strings.CreateAlertContentTypeAlertDescription,
      },
      {
        value: ContentType.Template,
        label: strings.CreateAlertContentTypeTemplateDescription,
      },
    ],
    [],
  );

  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );
  const { confirm, dialogs } = useFluentDialogs();

  const autoSaveStatusDebounceRef = React.useRef<number | null>(null);

  const copilotService = React.useMemo(
    () => new CopilotService(graphClient),
    [graphClient],
  );

  React.useEffect(() => {
    let isMounted = true;

    if (!copilotEnabled) {
      dispatch({
        type: "SET_COPILOT_AVAILABILITY",
        availability: "unknown",
      });
      return () => {
        isMounted = false;
      };
    }

    dispatch({
      type: "SET_COPILOT_AVAILABILITY",
      availability: "unknown",
    });

    copilotService
      .checkAccess()
      .then((isAvailable) => {
        if (!isMounted) {
          return;
        }

        dispatch({
          type: "SET_COPILOT_AVAILABILITY",
          availability: isAvailable ? "available" : "unavailable",
        });
      })
      .catch((error) => {
        logger.warn("CreateAlertTab", "Copilot access preflight failed", error);
        if (!isMounted) {
          return;
        }
        dispatch({
          type: "SET_COPILOT_AVAILABILITY",
          availability: "unavailable",
        });
      });

    return () => {
      isMounted = false;
      copilotService.cancelActiveOperation();
    };
  }, [copilotEnabled, copilotService, dispatch]);

  const languageOptions = useLanguageOptions(supportedLanguages);

  const loadSupportedLanguages = React.useCallback(async () => {
    try {
      const baseLanguages = LanguageAwarenessService.getSupportedLanguages();
      const supportedLanguageCodes = await alertService.getSupportedLanguages();

      logger.info(
        "CreateAlertTab",
        `Available language columns: ${supportedLanguageCodes.length}`,
      );

      const updatedLanguages = baseLanguages.map((lang) => ({
        ...lang,
        columnExists:
          supportedLanguageCodes.includes(lang.code) ||
          lang.code === TargetLanguage.EnglishUS,
        isSupported:
          supportedLanguageCodes.includes(lang.code) ||
          lang.code === TargetLanguage.EnglishUS,
      }));

      dispatch({
        type: "SET_SUPPORTED_LANGUAGES",
        languages: updatedLanguages,
      });
      logger.debug("CreateAlertTab", "Updated supported languages", {
        supportedLanguages: updatedLanguages
          .filter((l) => l.isSupported)
          .map((l) => l.code),
      });
    } catch (error) {
      logger.error(
        "CreateAlertTab",
        "Error loading supported languages",
        error,
      );
      const defaultLanguages = LanguageAwarenessService.getSupportedLanguages();
      dispatch({
        type: "SET_SUPPORTED_LANGUAGES",
        languages: defaultLanguages.map((lang) => ({
          ...lang,
          isSupported: lang.code === TargetLanguage.EnglishUS,
          columnExists: lang.code === TargetLanguage.EnglishUS,
        })),
      });
    }
  }, [alertService, dispatch]);

  React.useEffect(() => {
    loadSupportedLanguages();
  }, [loadSupportedLanguages]);

  React.useEffect(() => {
    let isMounted = true;
    alertService
      .getLanguagePolicy()
      .then((policy) => {
        if (isMounted) {
          dispatch({ type: "SET_LANGUAGE_POLICY", policy });
        }
      })
      .catch((error) => {
        logger.warn("CreateAlertTab", "Failed to load language policy", error);
      });
    return () => {
      isMounted = false;
    };
  }, [alertService, dispatch]);

  const loadDrafts = React.useCallback(async () => {
    try {
      const siteId = alertService.getCurrentSiteId();
      const userDrafts = await alertService.getDraftAlerts(siteId);
      dispatch({ type: "SET_DRAFTS", drafts: userDrafts });
    } catch (error) {
      logger.warn("CreateAlertTab", "Could not load drafts", error);
      dispatch({ type: "SET_DRAFTS", drafts: [] });
    }
  }, [alertService, dispatch]);

  React.useEffect(() => {
    loadDrafts();
  }, [loadDrafts]);

  const loadPreviousAlerts = React.useCallback(async () => {
    try {
      const currentSiteId = alertService.getCurrentSiteId();
      const existingItems = await alertService.getAlertsAndTemplates([
        currentSiteId,
      ]);
      const userDisplayName = (context.pageContext.user.displayName || "")
        .trim()
        .toLowerCase();
      const userEmail = (context.pageContext.user.email || "")
        .trim()
        .toLowerCase();
      const userLoginName = (
        (context.pageContext.user as unknown as { loginName?: string })
          ?.loginName || ""
      )
        .trim()
        .toLowerCase();

      const authorCandidates = new Set(
        [userDisplayName, userEmail, userLoginName].filter(Boolean),
      );

      const now = Date.now();
      const recentAlerts = existingItems
        .filter((item) => item.contentType === ContentType.Alert)
        .filter((item) => {
          const endValue = item.scheduledEnd
            ? new Date(item.scheduledEnd)
            : null;
          return (
            !!endValue &&
            !Number.isNaN(endValue.getTime()) &&
            endValue.getTime() < now
          );
        })
        .filter((item) => {
          const author = (item.createdBy || "").trim().toLowerCase();
          if (!author) {
            return false;
          }

          if (authorCandidates.has(author)) {
            return true;
          }

          for (const candidate of authorCandidates) {
            if (author.includes(candidate)) {
              return true;
            }
          }

          return false;
        })
        .sort((a, b) => {
          const aTime = a.scheduledEnd ? new Date(a.scheduledEnd).getTime() : 0;
          const bTime = b.scheduledEnd ? new Date(b.scheduledEnd).getTime() : 0;
          return bTime - aTime;
        })
        .slice(0, 12);
      dispatch({ type: "SET_PREVIOUS_ALERTS", alerts: recentAlerts });
    } catch (error) {
      logger.warn("CreateAlertTab", "Could not load previous alerts", error);
      dispatch({ type: "SET_PREVIOUS_ALERTS", alerts: [] });
    }
  }, [alertService, context.pageContext.user, dispatch]);

  React.useEffect(() => {
    loadPreviousAlerts();
  }, [loadPreviousAlerts]);

  const handleLoadDraft = React.useCallback(
    (draft: IAlertItem) => {
      dispatch({ type: "LOAD_DRAFT", draft });
    },
    [dispatch],
  );

  const handleDeleteDraft = React.useCallback(
    async (draft: IAlertItem) => {
      const shouldDelete = await confirm({
        title: strings.ManageAlertsDeleteDialogTitle,
        message: Text.format(
          strings.ManageAlertsDeleteDialogMessage,
          draft.title || strings.ManageAlertsNoTitleFallback,
        ),
        confirmText: strings.Delete,
        cancelText: strings.Cancel,
      });

      if (!shouldDelete) {
        return;
      }

      dispatch({ type: "SET_IS_CREATING", isCreating: true });
      try {
        await alertService.deleteDraft(draft.id);
        dispatch({
          type: "SET_DRAFTS",
          drafts: drafts.filter((item) => item.id !== draft.id),
        });

        if (autoSaveDraftId === draft.id) {
          dispatch({ type: "SET_AUTO_SAVE_DRAFT_ID", id: null });
        }

        notificationService.showSuccess(
          Text.format(
            strings.ManageAlertsDeleteSuccessMessage,
            draft.title || strings.ManageAlertsNoTitleFallback,
          ),
          strings.ManageAlertsDeleteSuccessTitle,
        );
      } catch (error) {
        logger.error("CreateAlertTab", "Error deleting draft", error);
        notificationService.showError(
          strings.ManageAlertsDeleteErrorMessage,
          strings.ManageAlertsDeleteErrorTitle,
        );
      } finally {
        dispatch({ type: "SET_IS_CREATING", isCreating: false });
      }
    },
    [
      alertService,
      autoSaveDraftId,
      confirm,
      notificationService,
      drafts,
      dispatch,
    ],
  );

  const handleLoadPreviousAlert = React.useCallback(
    (alert: IAlertItem) => {
      dispatch({ type: "LOAD_PREVIOUS_ALERT", alert });
    },
    [dispatch],
  );

  React.useEffect(() => {
    if (useMultiLanguage && newAlert.languageContent.length === 0) {
      dispatch({
        type: "SET_ALERT",
        payload: {
          languageContent: [
            {
              language: TargetLanguage.EnglishUS,
              title: newAlert.title,
              description: newAlert.description,
              linkDescription: newAlert.linkUrl
                ? newAlert.linkDescription
                : undefined,
              translationStatus: languagePolicy.workflow.enabled
                ? languagePolicy.workflow.defaultStatus
                : TranslationStatus.Approved,
            },
          ],
          languageGroup:
            newAlert.languageGroup ||
            `lg-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
        },
      });
    } else if (!useMultiLanguage && newAlert.languageContent.length > 0) {
      const firstLang = newAlert.languageContent[0];
      if (firstLang) {
        dispatch({
          type: "SET_ALERT",
          payload: {
            title: firstLang.title,
            description: firstLang.description,
            linkDescription: firstLang.linkDescription || "",
            languageContent: [],
            languageGroup: undefined,
          },
        });
      }
    }
  }, [useMultiLanguage, newAlert.languageContent.length, dispatch]);

  const handleTemplateSelect = React.useCallback(
    async (template: IAlertTemplate) => {
      try {
        const templateAlerts = await alertService.getTemplateAlerts(
          alertService.getCurrentSiteId(),
        );
        const fullTemplate = templateAlerts.find(
          (t: IAlertItem) => t.id === template.id,
        );

        if (fullTemplate) {
          dispatch({
            type: "LOAD_TEMPLATE",
            template: fullTemplate,
            useMultiLanguage,
          });

          if (useMultiLanguage && fullTemplate.languageGroup) {
            try {
              const languageVariants = templateAlerts.filter(
                (a: IAlertItem) =>
                  a.languageGroup === fullTemplate.languageGroup,
              );

              const languageContent = languageVariants.map(
                (variant: IAlertItem) => ({
                  language: variant.targetLanguage,
                  title: variant.title,
                  description: variant.description,
                  linkDescription: variant.linkDescription || "",
                }),
              );

              if (languageContent.length > 0) {
                dispatch({
                  type: "SET_ALERT",
                  payload: {
                    languageContent,
                    targetLanguage: languageContent[0].language,
                  },
                });
              }
            } catch (error) {
              logger.warn(
                "CreateAlertTab",
                "Could not load template language variants",
                error,
              );
            }
          }
        } else {
          dispatch({
            type: "SET_ALERT",
            payload: {
              title: template.template.title,
              description: template.template.description,
              priority: template.template.priority,
              notificationType: template.template.notificationType,
              isPinned: template.template.isPinned,
              linkUrl: template.template.linkUrl || "",
              linkDescription: template.template.linkDescription || "",
              contentType: ContentType.Alert,
            },
          });
        }
      } catch (error) {
        logger.error("CreateAlertTab", "Failed to load template", error);
        dispatch({
          type: "SET_ALERT",
          payload: {
            title: template.template.title,
            description: template.template.description,
            priority: template.template.priority,
            notificationType: template.template.notificationType,
            isPinned: template.template.isPinned,
            linkUrl: template.template.linkUrl || "",
            linkDescription: template.template.linkDescription || "",
            contentType: ContentType.Alert,
          },
        });
      }

      dispatch({ type: "SET_ENTRY_MODE", mode: "scratch" });
      dispatch({ type: "SET_CREATE_STEP", step: "content" });
      dispatch({ type: "CLEAR_ERRORS" });
    },
    [alertService, dispatch, useMultiLanguage],
  );

  const validationSnapshot = React.useMemo(() => {
    const validationData = {
      ...newAlert,
      scheduledStart: newAlert.scheduledStart
        ? newAlert.scheduledStart.toISOString()
        : undefined,
      scheduledEnd: newAlert.scheduledEnd
        ? newAlert.scheduledEnd.toISOString()
        : undefined,
    };
    return validateAlertData(validationData, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage: languageService.getTenantDefaultLanguage(),
      getString: getLocalizedValidationMessage,
      validateTargetSites: enableTargetSite,
    });
  }, [
    enableTargetSite,
    languagePolicy,
    languageService,
    newAlert,
    useMultiLanguage,
  ]);

  const stepCompletion = React.useMemo<Record<CreateWizardStep, boolean>>(
    () => ({
      content:
        !validationSnapshot.title &&
        !validationSnapshot.description &&
        !validationSnapshot.AlertType &&
        !validationSnapshot.languageContent,
      audience:
        !validationSnapshot.linkUrl &&
        !validationSnapshot.linkDescription &&
        (!enableTargetSite || !validationSnapshot.targetSites),
      publish:
        !validationSnapshot.scheduledStart && !validationSnapshot.scheduledEnd,
    }),
    [enableTargetSite, validationSnapshot],
  );

  const getStepErrors = React.useCallback(
    (step: CreateWizardStep): Record<string, string> => {
      const fieldByStep: Record<CreateWizardStep, string[]> = {
        content: ["title", "description", "AlertType", "languageContent"],
        audience: ["linkUrl", "linkDescription", "targetSites"],
        publish: ["scheduledStart", "scheduledEnd"],
      };

      const stepErrors: Record<string, string> = {};
      fieldByStep[step].forEach((field) => {
        const errorMessage = validationSnapshot[field];
        if (errorMessage) {
          stepErrors[field] = errorMessage;
        }
      });

      return stepErrors;
    },
    [validationSnapshot],
  );

  const handleCreateStepChange = React.useCallback(
    (nextStep: CreateWizardStep) => {
      if (createStep === nextStep) {
        return;
      }

      const currentStepErrors = getStepErrors(createStep);
      if (Object.keys(currentStepErrors).length > 0) {
        dispatch({ type: "SET_ERRORS", errors: currentStepErrors });
        return;
      }

      dispatch({ type: "SET_CREATE_STEP", step: nextStep });
    },
    [createStep, getStepErrors, dispatch],
  );

  const validateForm = React.useCallback((): boolean => {
    const validationData = {
      ...newAlert,
      scheduledStart: newAlert.scheduledStart
        ? newAlert.scheduledStart.toISOString()
        : undefined,
      scheduledEnd: newAlert.scheduledEnd
        ? newAlert.scheduledEnd.toISOString()
        : undefined,
    };

    const validationErrors = validateAlertData(validationData, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage: languageService.getTenantDefaultLanguage(),
      getString: getLocalizedValidationMessage,
      validateTargetSites: enableTargetSite,
    });

    dispatch({ type: "SET_ERRORS", errors: validationErrors });
    return Object.keys(validationErrors).length === 0;
  }, [
    newAlert,
    useMultiLanguage,
    languagePolicy,
    languageService,
    enableTargetSite,
    dispatch,
  ]);

  const handleCreateAlert = React.useCallback(async () => {
    if (!validateForm()) return;

    const summaryLines = [
      `${strings.AlertTitle}: ${newAlert.title || strings.ManageAlertsNoTitleFallback}`,
      `${strings.MultiLanguageContent}: ${
        useMultiLanguage
          ? Text.format(
              strings.MultiLanguageEditorLanguageCount,
              newAlert.languageContent.length.toString(),
            )
          : strings.CreateAlertSingleLanguageButton
      }`,
      `${strings.TargetSites}: ${newAlert.targetSites.length.toString()}`,
      `${strings.NotificationType}: ${newAlert.notificationType}`,
      `${strings.CreateAlertSchedulingSectionTitle}: ${Intl.DateTimeFormat().resolvedOptions().timeZone}`,
    ];

    const shouldCreate = await confirm({
      title: strings.CreateAlertPublishSummaryTitle,
      message: summaryLines.join("\n"),
      confirmText: strings.Create,
      cancelText: strings.Cancel,
    });

    if (!shouldCreate) {
      return;
    }

    dispatch({ type: "SET_IS_CREATING", isCreating: true });
    dispatch({ type: "SET_CREATION_PROGRESS", progress: [] });

    try {
      const createdVariants = await AlertMutationService.createAlertsFromModel(
        alertService,
        newAlert,
        {
          useMultiLanguage,
          createdBy: context.pageContext.user.displayName,
          languagePolicy,
          languageService,
        },
      );

      if (useMultiLanguage && newAlert.languageContent.length > 0) {
        dispatch({
          type: "SET_CREATION_PROGRESS",
          progress: [
            {
              siteId: "success",
              siteName: Text.format(
                strings.CreateAlertMultiLanguageSuccessStatus,
                createdVariants,
              ),
              hasAccess: true,
              canCreateAlerts: true,
              permissionLevel: "success",
              error: "",
            },
          ],
        });
        notificationService.showSuccess(
          Text.format(
            strings.CreateAlertMultiLanguageSuccessStatus,
            createdVariants,
          ),
          strings.CreateAlertMultiLanguageSuccessTitle,
        );
      } else {
        dispatch({
          type: "SET_CREATION_PROGRESS",
          progress: [
            {
              siteId: "success",
              siteName: strings.CreateAlertCreationSuccessStatus,
              hasAccess: true,
              canCreateAlerts: true,
              permissionLevel: "success",
              error: "",
            },
          ],
        });
        notificationService.showSuccess(
          strings.CreateAlertCreationSuccessStatus,
          strings.CreateAlertCreationSuccessTitle,
        );
      }

      dispatch({ type: "RESET_FORM" });
      await loadPreviousAlerts();
    } catch (error) {
      logger.error("CreateAlertTab", "Error creating alert", error);
      dispatch({ type: "SET_LAST_CREATE_FAILED", failed: true });
      dispatch({
        type: "SET_CREATION_PROGRESS",
        progress: [
          {
            siteId: "error",
            siteName: strings.CreateAlertCreationErrorStatus,
            hasAccess: false,
            canCreateAlerts: false,
            permissionLevel: "error",
            error:
              error instanceof Error
                ? error.message
                : strings.CreateAlertUnknownError,
          },
        ],
      });
      notificationService.showError(
        strings.CreateAlertCreationErrorStatus,
        strings.CreateAlertCreationErrorTitle,
      );
    } finally {
      dispatch({ type: "SET_IS_CREATING", isCreating: false });
    }
  }, [
    alertService,
    confirm,
    context.pageContext.user.displayName,
    languagePolicy,
    languageService,
    loadPreviousAlerts,
    newAlert,
    notificationService,
    useMultiLanguage,
    validateForm,
    dispatch,
  ]);

  const handleSaveAsDraft = React.useCallback(async () => {
    if (!newAlert.title || newAlert.title.trim().length < 3) {
      dispatch({
        type: "SET_ERRORS",
        errors: { title: strings.CreateAlertDraftTitleError },
      });
      return;
    }

    if (
      !newAlert.AlertType ||
      !alertTypes.find((t) => t.name === newAlert.AlertType)
    ) {
      dispatch({
        type: "SET_ERRORS",
        errors: { AlertType: strings.CreateAlertAlertTypeError },
      });
      return;
    }

    dispatch({ type: "SET_IS_CREATING", isCreating: true });
    try {
      const draftData = AlertMutationService.buildDraftPayload(newAlert, {
        createdBy: context.pageContext.user.displayName,
      });
      await alertService.saveDraft(draftData);

      dispatch({
        type: "SET_CREATION_PROGRESS",
        progress: [
          {
            siteId: "success",
            siteName: strings.CreateAlertDraftSaveSuccessStatus,
            hasAccess: true,
            canCreateAlerts: true,
            permissionLevel: "success",
            error: "",
          },
        ],
      });
      dispatch({ type: "SET_LAST_CREATE_FAILED", failed: false });

      await loadDrafts();
      setTimeout(() => {
        resetForm();
        dispatch({ type: "SET_CREATION_PROGRESS", progress: [] });
      }, 2000);
    } catch (error) {
      logger.error("CreateAlertTab", "Error saving draft", error);
      dispatch({
        type: "SET_CREATION_PROGRESS",
        progress: [
          {
            siteId: "error",
            siteName: strings.CreateAlertDraftSaveErrorStatus,
            hasAccess: false,
            canCreateAlerts: false,
            permissionLevel: "error",
            error:
              error instanceof Error
                ? error.message
                : strings.CreateAlertUnknownError,
          },
        ],
      });
      dispatch({ type: "SET_LAST_CREATE_FAILED", failed: true });
    } finally {
      dispatch({ type: "SET_IS_CREATING", isCreating: false });
    }
  }, [
    newAlert,
    alertService,
    context.pageContext.user.displayName,
    alertTypes,
    loadDrafts,
    dispatch,
  ]);

  const autoSaveDraft = React.useCallback(async () => {
    if (
      !newAlert.title ||
      newAlert.title.trim().length < 3 ||
      !newAlert.AlertType
    ) {
      return;
    }

    if (isCreatingAlert || isAutoSaving) {
      return;
    }

    dispatch({ type: "SET_AUTO_SAVE_STATUS", status: "saving" });
    dispatch({ type: "SET_IS_AUTO_SAVING", isSaving: true });
    try {
      const draftData = AlertMutationService.buildDraftPayload(newAlert, {
        id: autoSaveDraftId || undefined,
        createdBy: context.pageContext.user.displayName,
        titlePrefix: "[Auto-saved] ",
      });
      const savedDraft = await alertService.saveDraft(draftData);
      dispatch({ type: "SET_AUTO_SAVE_DRAFT_ID", id: savedDraft.id });
      dispatch({ type: "SET_LAST_AUTO_SAVE", date: new Date() });
      dispatch({ type: "SET_AUTO_SAVE_STATUS", status: "saved" });
      logger.debug("CreateAlertTab", "Auto-saved draft", {
        id: savedDraft.id,
        title: savedDraft.title,
      });
    } catch (error) {
      logger.warn("CreateAlertTab", "Auto-save failed", error);
      dispatch({ type: "SET_AUTO_SAVE_STATUS", status: "error" });
    } finally {
      dispatch({ type: "SET_IS_AUTO_SAVING", isSaving: false });
    }
  }, [
    newAlert,
    isCreatingAlert,
    isAutoSaving,
    autoSaveDraftId,
    alertService,
    context.pageContext.user.displayName,
    dispatch,
  ]);

  React.useEffect(() => {
    const hasDraftableContent =
      newAlert.title.trim().length > 0 ||
      newAlert.description.trim().length > 0 ||
      newAlert.linkUrl.trim().length > 0;

    if (!hasDraftableContent) {
      dispatch({ type: "SET_AUTO_SAVE_STATUS", status: "idle" });
      return;
    }

    if (isCreatingAlert || isAutoSaving) {
      return;
    }

    if (autoSaveStatusDebounceRef.current) {
      window.clearTimeout(autoSaveStatusDebounceRef.current);
    }

    autoSaveStatusDebounceRef.current = window.setTimeout(() => {
      dispatch({ type: "SET_AUTO_SAVE_STATUS", status: "pending" });
    }, 800);

    return () => {
      if (autoSaveStatusDebounceRef.current) {
        window.clearTimeout(autoSaveStatusDebounceRef.current);
      }
    };
  }, [newAlert, isCreatingAlert, isAutoSaving, dispatch]);

  React.useEffect(() => {
    const intervalId = setInterval(autoSaveDraft, 30000);
    return () => clearInterval(intervalId);
  }, [autoSaveDraft]);

  const resetForm = React.useCallback(() => {
    dispatch({ type: "RESET_FORM" });
  }, [dispatch]);

  React.useEffect(() => {
    if (currentEntryMode !== "scratch") {
      return;
    }

    const handler = (event: KeyboardEvent): void => {
      if (!(event.ctrlKey || event.metaKey)) {
        return;
      }

      if (event.key.toLowerCase() === "s") {
        event.preventDefault();
        if (!isCreatingAlert) {
          void handleSaveAsDraft();
        }
      }

      if (event.key === "Enter" && createStep === "publish") {
        event.preventDefault();
        if (!isCreatingAlert) {
          void handleCreateAlert();
        }
      }
    };

    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [
    createStep,
    currentEntryMode,
    handleCreateAlert,
    handleSaveAsDraft,
    isCreatingAlert,
  ]);

  const stepOrder: CreateWizardStep[] = ["content", "audience", "publish"];
  const createStepIndex = stepOrder.indexOf(createStep);
  const canGoBack = createStepIndex > 0;
  const canGoForward = createStepIndex < stepOrder.length - 1;

  const handleBackStep = React.useCallback(() => {
    if (!canGoBack) {
      return;
    }
    dispatch({ type: "SET_CREATE_STEP", step: stepOrder[createStepIndex - 1] });
  }, [canGoBack, createStepIndex, stepOrder, dispatch]);

  const handleNextStep = React.useCallback(() => {
    if (!canGoForward) {
      return;
    }
    handleCreateStepChange(stepOrder[createStepIndex + 1]);
  }, [canGoForward, createStepIndex, handleCreateStepChange, stepOrder]);

  const imageFolderName =
    newAlert.languageGroup ||
    StringUtils.sanitizeForId(newAlert.title) ||
    "Untitled_Alert";

  const setEntryMode = React.useCallback(
    (mode: "scratch" | "templates" | "drafts" | "previous") => {
      dispatch({ type: "SET_ENTRY_MODE", mode });
    },
    [dispatch],
  );

  return (
    <div className={styles.tabPane}>
      {dialogs}

      <div className={styles.templateActions}>
        <SharePointButton
          variant={currentEntryMode === "scratch" ? "primary" : "secondary"}
          onClick={() => setEntryMode("scratch")}
        >
          {strings.CreateAlertStartFromScratchButton}
        </SharePointButton>
        <SharePointButton
          variant={currentEntryMode === "templates" ? "primary" : "secondary"}
          onClick={() => setEntryMode("templates")}
        >
          {strings.CreateAlertTemplatesButtonLabel}
        </SharePointButton>
        <SharePointButton
          variant={currentEntryMode === "previous" ? "primary" : "secondary"}
          onClick={() => setEntryMode("previous")}
          icon={<History24Regular />}
        >
          {Text.format(
            strings.CreateAlertPreviousAlertsButtonLabel,
            previousAlerts.length,
          )}
        </SharePointButton>
        <SharePointButton
          variant={currentEntryMode === "drafts" ? "primary" : "secondary"}
          onClick={() => setEntryMode("drafts")}
        >
          {Text.format(strings.CreateAlertDraftsButtonLabel, drafts.length)}
        </SharePointButton>
      </div>

      {currentEntryMode === "templates" && (
        <div className={styles.templatesSection}>
          <React.Suspense
            fallback={
              <div className={styles.emptyState}>
                <p>{strings.Loading}</p>
              </div>
            }
          >
            <AlertTemplates
              onSelectTemplate={handleTemplateSelect}
              graphClient={graphClient}
              context={context}
              alertService={alertService}
              className={styles.templates}
            />
          </React.Suspense>
        </div>
      )}

      {currentEntryMode === "previous" && (
        <div className={styles.templatesSection}>
          <div className={styles.alertsList}>
            {previousAlerts.length === 0 ? (
              <EmptyState
                icon="History"
                title={strings.CreateAlertNoPreviousAlertsTitle}
                description={strings.CreateAlertPreviousAlertsEmptyMessage}
              />
            ) : (
              previousAlerts.map((alert) => (
                <div key={alert.id} className={styles.alertCard}>
                  <div className={styles.alertCardHeader}>
                    <h4>{alert.title}</h4>
                    <small>
                      {DateUtils.formatForDisplay(alert.createdDate)}
                    </small>
                  </div>
                  <div className={styles.alertCardContent}>
                    <div
                      dangerouslySetInnerHTML={{
                        __html: StringUtils.truncate(alert.description, 140),
                      }}
                    />
                  </div>
                  <div className={styles.alertCardActions}>
                    <SharePointButton
                      variant="primary"
                      onClick={() => handleLoadPreviousAlert(alert)}
                      icon={<DocumentArrowLeft24Regular />}
                    >
                      {strings.CreateAlertUsePreviousButton}
                    </SharePointButton>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {currentEntryMode === "drafts" && (
        <div className={styles.templatesSection}>
          <div className={styles.alertsList}>
            {drafts.length === 0 ? (
              <EmptyState
                icon="Draft"
                title={strings.CreateAlertNoDraftsTitle}
                description={strings.CreateAlertDraftsEmptyMessage}
              />
            ) : (
              drafts.map((draft) => (
                <div key={draft.id} className={styles.alertCard}>
                  <div className={styles.alertCardHeader}>
                    <h4>{draft.title}</h4>
                    <small>
                      {DateUtils.formatForDisplay(draft.createdDate)}
                    </small>
                  </div>
                  <div className={styles.alertCardContent}>
                    <div
                      dangerouslySetInnerHTML={{
                        __html: StringUtils.truncate(draft.description, 100),
                      }}
                    />
                  </div>
                  <div className={styles.alertCardActions}>
                    <SharePointButton
                      variant="primary"
                      onClick={() => handleLoadDraft(draft)}
                      icon={<DocumentArrowLeft24Regular />}
                    >
                      {strings.CreateAlertLoadDraftButton}
                    </SharePointButton>
                    <SharePointButton
                      variant="danger"
                      onClick={() => void handleDeleteDraft(draft)}
                      icon={<Delete24Regular />}
                    >
                      {strings.Delete}
                    </SharePointButton>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {currentEntryMode === "scratch" && (
        <AlertEditorForm
          mode="create"
          alert={newAlert}
          setAlert={(value) => {
            if (typeof value === "function") {
              const prev = newAlert;
              const next = value(
                prev as INewAlert & { contentType: ContentType },
              );
              dispatch({ type: "SET_ALERT", payload: next });
            } else {
              dispatch({ type: "SET_ALERT", payload: value });
            }
          }}
          errors={errors}
          setErrors={(errorsOrFn) => {
            if (typeof errorsOrFn === "function") {
              const next = errorsOrFn(errors);
              dispatch({ type: "SET_ERRORS", errors: next });
            } else {
              dispatch({ type: "SET_ERRORS", errors: errorsOrFn });
            }
          }}
          alertTypes={alertTypes}
          alertTypeOptions={alertTypeOptions}
          notificationOptions={notificationOptions}
          contentTypeOptions={contentTypeOptions}
          languageOptions={languageOptions}
          supportedLanguages={supportedLanguages}
          useMultiLanguage={useMultiLanguage}
          setUseMultiLanguage={(
            use: boolean | ((prev: boolean) => boolean),
          ) => {
            const value =
              typeof use === "function" ? use(useMultiLanguage) : use;
            dispatch({
              type: "SET_USE_MULTI_LANGUAGE",
              useMultiLanguage: value,
            });
          }}
          userTargetingEnabled={userTargetingEnabled}
          notificationsEnabled={notificationsEnabled}
          enableTargetSite={enableTargetSite}
          siteDetector={siteDetector}
          graphClient={graphClient}
          context={context}
          languagePolicy={languagePolicy}
          tenantDefaultLanguage={languageService.getTenantDefaultLanguage()}
          copilotEnabled={copilotEnabled}
          copilotService={copilotService}
          notificationService={notificationService}
          isBusy={isCreatingAlert}
          showPreview={showPreview}
          setShowPreview={(show: boolean | ((prev: boolean) => boolean)) => {
            const value = typeof show === "function" ? show(showPreview) : show;
            dispatch({ type: "SET_SHOW_PREVIEW", show: value });
          }}
          imageFolderName={imageFolderName}
          createStep={createStep}
          onCreateStepChange={handleCreateStepChange}
          stepCompletion={stepCompletion}
          copilotAvailability={copilotAvailability}

          actionButtons={
            <>
              <SharePointButton
                variant="secondary"
                onClick={handleBackStep}
                disabled={!canGoBack || isCreatingAlert}
                icon={<ArrowLeft24Regular />}
              >
                {strings.CreateAlertWizardBackButton}
              </SharePointButton>

              {canGoForward ? (
                <SharePointButton
                  variant="primary"
                  onClick={handleNextStep}
                  disabled={isCreatingAlert}
                  icon={<ArrowRight24Regular />}
                >
                  {strings.CreateAlertWizardNextButton}
                </SharePointButton>
              ) : (
                <SharePointButton
                  variant="primary"
                  onClick={() => void handleCreateAlert()}
                  disabled={isCreatingAlert || alertTypes.length === 0}
                  icon={<Save24Regular />}
                >
                  {isCreatingAlert
                    ? strings.CreateAlertPrimaryButtonLoading
                    : strings.CreateAlert}
                </SharePointButton>
              )}

              <SharePointButton
                variant="secondary"
                onClick={() => void handleSaveAsDraft()}
                disabled={isCreatingAlert || alertTypes.length === 0}
                icon={<Drafts24Regular />}
              >
                {strings.CreateAlertSaveDraftButtonLabel}
              </SharePointButton>

              <SharePointButton
                variant="secondary"
                onClick={resetForm}
                disabled={isCreatingAlert}
                icon={<Dismiss24Regular />}
              >
                {strings.CreateAlertResetFormButtonLabel}
              </SharePointButton>
            </>
          }
          afterActions={
            <>
              {autoSaveStatus !== "idle" && (
                <div className={styles.autoSaveIndicator}>
                  {autoSaveStatus === "saving" && (
                    <span>{strings.CreateAlertAutoSaving}</span>
                  )}
                  {autoSaveStatus === "pending" && (
                    <span>{strings.CreateAlertAutoSavePending}</span>
                  )}
                  {autoSaveStatus === "error" && (
                    <span>{strings.CreateAlertAutoSaveRetrying}</span>
                  )}
                  {autoSaveStatus === "saved" && lastAutoSave && (
                    <span>
                      {Text.format(
                        strings.CreateAlertAutoSavedAt,
                        lastAutoSave.toLocaleTimeString(),
                      )}
                    </span>
                  )}
                </div>
              )}
              {creationProgress.length > 0 && (
                <div className={styles.alertsList}>
                  <h3>{strings.CreateAlertCreationResultsHeading}</h3>
                  {creationProgress.map((result, index) => (
                    <div
                      key={index}
                      className={`${styles.alertCard} ${result.error ? styles.error : styles.success}`}
                    >
                      <strong>{result.siteName}</strong>:{" "}
                      {result.error
                        ? Text.format(
                            strings.CreateAlertCreationResultError,
                            result.error,
                          )
                        : strings.CreateAlertCreationResultSuccess}
                    </div>
                  ))}
                  {lastCreateAttemptFailed && (
                    <div className={styles.retryAction}>
                      <SharePointButton
                        variant="secondary"
                        onClick={() => void handleCreateAlert()}
                      >
                        {strings.CreateAlertCreationRetryButton}
                      </SharePointButton>
                    </div>
                  )}
                </div>
              )}
            </>
          }
        />
      )}
    </div>
  );
};


/**
 * Wrapper component that extracts props and provides them to the inner component
 * This maintains backward compatibility with existing code while using context internally
 */
const CreateAlertTab: React.FC<ICreateAlertTabProps> = (props) => {
  useAlertFormState();

  return <CreateAlertTabInner {...props} />;
};

export default CreateAlertTab;
