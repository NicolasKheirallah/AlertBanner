import * as React from "react";
import {
  Save24Regular,
  Dismiss24Regular,
  Drafts24Regular,
  DocumentArrowLeft24Regular,
  ArrowLeft24Regular,
  ArrowRight24Regular,
  History24Regular,
} from "@fluentui/react-icons";
import {
  SharePointButton,
  ISharePointSelectOption,
} from "../../UI/SharePointControls";
import type { IAlertTemplate } from "../../UI/AlertTemplates";
import {
  AlertPriority,
  NotificationType,
  IAlertType,
  ContentType,
  TargetLanguage,
  IPersonField,
  IAlertItem,
  TranslationStatus,
} from "../../Alerts/IAlerts";
import {
  LanguageAwarenessService,
  ILanguageContent,
  ISupportedLanguage,
} from "../../Services/LanguageAwarenessService";
import {
  DEFAULT_LANGUAGE_POLICY,
  ILanguagePolicy,
} from "../../Services/LanguagePolicyService";
import {
  SiteContextDetector,
  ISiteValidationResult,
} from "../../Utils/SiteContextDetector";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { logger } from "../../Services/LoggerService";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { NotificationService } from "../../Services/NotificationService";
import { CopilotService } from "../../Services/CopilotService";
import styles from "../AlertSettings.module.scss";
import {
  validateAlertData,
} from "../../Utils/AlertValidation";
import { getLocalizedValidationMessage } from "../../Utils/AlertValidationLocalization";
import { useLanguageOptions } from "../../Hooks/useLanguageOptions";
import { usePriorityOptions } from "../../Hooks/usePriorityOptions";
import { DateUtils } from "../../Utils/DateUtils";
import { StringUtils } from "../../Utils/StringUtils";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import AlertEditorForm, { CreateWizardStep } from "./AlertEditorForm";
import { AlertMutationService } from "../../Services/AlertMutationService";
import { useFluentDialogs } from "../../Hooks/useFluentDialogs";

const AlertTemplates = React.lazy(() => import("../../UI/AlertTemplates"));

type AutoSaveStatus = "idle" | "pending" | "saving" | "saved" | "error";
type CreateEntryMode = "scratch" | "templates" | "drafts" | "previous";

const stripHtml = (value: string): string =>
  value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();

export interface INewAlert {
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl: string;
  linkDescription: string;
  targetSites: string[];
  scheduledStart?: Date;
  scheduledEnd?: Date;
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageContent: ILanguageContent[];
  languageGroup?: string; // For multi-language alerts to share resources
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
}

import { IFormErrors } from "./SharedTypes";
export type { IFormErrors };

export interface ICreateAlertTabProps {
  newAlert: INewAlert;
  setNewAlert: React.Dispatch<React.SetStateAction<INewAlert>>;
  errors: IFormErrors;
  setErrors: React.Dispatch<React.SetStateAction<IFormErrors>>;
  alertTypes: IAlertType[];
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite?: boolean; // Optional - defaults to false
  siteDetector: SiteContextDetector;
  alertService: SharePointAlertService;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  creationProgress: ISiteValidationResult[];
  setCreationProgress: React.Dispatch<
    React.SetStateAction<ISiteValidationResult[]>
  >;
  isCreatingAlert: boolean;
  setIsCreatingAlert: React.Dispatch<React.SetStateAction<boolean>>;
  showPreview: boolean;
  setShowPreview: React.Dispatch<React.SetStateAction<boolean>>;
  showTemplates: boolean;
  setShowTemplates: React.Dispatch<React.SetStateAction<boolean>>;
  languageUpdateTrigger?: number;
  copilotEnabled?: boolean;
  onDirtyStateChange?: (hasUnsavedChanges: boolean) => void;
}

const CreateAlertTab: React.FC<ICreateAlertTabProps> = ({
  newAlert,
  setNewAlert,
  errors,
  setErrors,
  alertTypes,
  userTargetingEnabled,
  notificationsEnabled,
  enableTargetSite = false, // Default to false
  siteDetector,
  alertService,
  graphClient,
  context,
  creationProgress,
  setCreationProgress,
  isCreatingAlert,
  setIsCreatingAlert,
  showPreview,
  setShowPreview,
  showTemplates,
  setShowTemplates,
  languageUpdateTrigger,
  copilotEnabled,
  onDirtyStateChange,
}) => {
  // Priority options - using shared hook
  const priorityOptions = usePriorityOptions();

  // Notification type options with detailed descriptions
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

  // Alert type options
  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(
    (type) => ({
      value: type.name,
      label: type.name,
    }),
  );

  // Content type options
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

  // Language awareness state
  const [languageService] = React.useState(
    () => new LanguageAwarenessService(graphClient, context),
  );
  const [supportedLanguages, setSupportedLanguages] = React.useState<
    ISupportedLanguage[]
  >([]);
  const [useMultiLanguage, setUseMultiLanguage] = React.useState(false);
  const [languagePolicy, setLanguagePolicy] = React.useState<ILanguagePolicy>(
    DEFAULT_LANGUAGE_POLICY,
  );
  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );
  const { confirm, dialogs } = useFluentDialogs();

  // Draft state
  const [drafts, setDrafts] = React.useState<IAlertItem[]>([]);
  const [previousAlerts, setPreviousAlerts] = React.useState<IAlertItem[]>([]);
  const [showDrafts, setShowDrafts] = React.useState(false);
  const [showPrevious, setShowPrevious] = React.useState(false);
  const [autoSaveDraftId, setAutoSaveDraftId] = React.useState<string | null>(
    null,
  );
  const [lastAutoSave, setLastAutoSave] = React.useState<Date | null>(null);
  const [isAutoSaving, setIsAutoSaving] = React.useState(false);
  const [autoSaveStatus, setAutoSaveStatus] =
    React.useState<AutoSaveStatus>("idle");
  const [createStep, setCreateStep] = React.useState<CreateWizardStep>("content");
  const [lastCreateAttemptFailed, setLastCreateAttemptFailed] =
    React.useState(false);
  const [copilotAvailability, setCopilotAvailability] = React.useState<
    "unknown" | "available" | "unavailable"
  >("unknown");
  const [copilotAvailabilityMessage, setCopilotAvailabilityMessage] =
    React.useState<string>("");
  const autoSaveStatusDebounceRef = React.useRef<number | null>(null);

  const currentEntryMode = React.useMemo<CreateEntryMode>(() => {
    if (showTemplates) {
      return "templates";
    }

    if (showPrevious) {
      return "previous";
    }

    if (showDrafts) {
      return "drafts";
    }

    return "scratch";
  }, [showDrafts, showPrevious, showTemplates]);

  const setEntryMode = React.useCallback(
    (mode: CreateEntryMode) => {
      setShowDrafts(mode === "drafts");
      setShowPrevious(mode === "previous");
      setShowTemplates(mode === "templates");
    },
    [setShowDrafts, setShowTemplates],
  );

  // Copilot state
  const copilotService = React.useMemo(
    () => new CopilotService(graphClient),
    [graphClient],
  );

  React.useEffect(() => {
    let isMounted = true;

    if (!copilotEnabled) {
      setCopilotAvailability("unknown");
      setCopilotAvailabilityMessage("");
      return () => {
        isMounted = false;
      };
    }

    setCopilotAvailability("unknown");
    setCopilotAvailabilityMessage("");

    copilotService
      .checkAccess()
      .then((isAvailable) => {
        if (!isMounted) {
          return;
        }

        setCopilotAvailability(isAvailable ? "available" : "unavailable");
        setCopilotAvailabilityMessage(
          isAvailable ? "" : strings.CreateAlertCopilotUnavailableMessage,
        );
      })
      .catch((error) => {
        logger.warn("CreateAlertTab", "Copilot access preflight failed", error);
        if (!isMounted) {
          return;
        }
        setCopilotAvailability("unavailable");
        setCopilotAvailabilityMessage(strings.CopilotUnexpectedError);
      });

    return () => {
      isMounted = false;
      copilotService.cancelActiveOperation();
    };
  }, [copilotEnabled, copilotService]);

  // Language targeting options - using shared hook
  const languageOptions = useLanguageOptions(supportedLanguages);

  // Load supported languages from SharePoint (actual enabled ones)
  const loadSupportedLanguages = React.useCallback(async () => {
    try {
      // Get the base language definitions
      const baseLanguages = LanguageAwarenessService.getSupportedLanguages();

      // Get the actually supported languages from SharePoint columns
      const supportedLanguageCodes = await alertService.getSupportedLanguages();

      logger.info(
        "CreateAlertTab",
        `Available language columns: ${supportedLanguageCodes.length}`,
      );

      // Update the base languages with the actual status
      const updatedLanguages = baseLanguages.map((lang) => ({
        ...lang,
        columnExists:
          supportedLanguageCodes.includes(lang.code) ||
          lang.code === TargetLanguage.EnglishUS,
        isSupported:
          supportedLanguageCodes.includes(lang.code) ||
          lang.code === TargetLanguage.EnglishUS,
      }));

      setSupportedLanguages(updatedLanguages);
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
      // Fallback to default with English only
      const defaultLanguages = LanguageAwarenessService.getSupportedLanguages();
      setSupportedLanguages(
        defaultLanguages.map((lang) => ({
          ...lang,
          isSupported: lang.code === TargetLanguage.EnglishUS,
          columnExists: lang.code === TargetLanguage.EnglishUS,
        })),
      );
    }
  }, [alertService]);

  React.useEffect(() => {
    loadSupportedLanguages();
  }, [loadSupportedLanguages, languageUpdateTrigger]);

  React.useEffect(() => {
    let isMounted = true;
    alertService
      .getLanguagePolicy()
      .then((policy) => {
        if (isMounted) {
          setLanguagePolicy(policy);
        }
      })
      .catch((error) => {
        logger.warn("CreateAlertTab", "Failed to load language policy", error);
      });
    return () => {
      isMounted = false;
    };
  }, [alertService]);

  // Load drafts
  const loadDrafts = React.useCallback(async () => {
    try {
      const siteId = alertService.getCurrentSiteId();
      const userDrafts = await alertService.getDraftAlerts(siteId);
      setDrafts(userDrafts);
    } catch (error) {
      logger.warn("CreateAlertTab", "Could not load drafts", error);
      setDrafts([]);
    }
  }, [alertService]);

  React.useEffect(() => {
    loadDrafts();
  }, [loadDrafts]);

  const loadPreviousAlerts = React.useCallback(async () => {
    try {
      const currentSiteId = alertService.getCurrentSiteId();
      const existingItems = await alertService.getAlertsAndTemplates([
        currentSiteId,
      ]);
      const recentAlerts = existingItems
        .filter(
          (item) =>
            item.contentType === ContentType.Alert &&
            item.contentStatus !== "Draft",
        )
        .slice(0, 12);
      setPreviousAlerts(recentAlerts);
    } catch (error) {
      logger.warn("CreateAlertTab", "Could not load previous alerts", error);
      setPreviousAlerts([]);
    }
  }, [alertService]);

  React.useEffect(() => {
    loadPreviousAlerts();
  }, [loadPreviousAlerts]);

  // Load from draft
  const handleLoadDraft = React.useCallback(
    (draft: IAlertItem) => {
      setNewAlert({
        title: draft.title,
        description: draft.description,
        AlertType: draft.AlertType,
        priority: draft.priority,
        isPinned: draft.isPinned,
        notificationType: draft.notificationType,
        linkUrl: draft.linkUrl || "",
        linkDescription: draft.linkDescription || "",
        targetSites: draft.targetSites || [],
        scheduledStart: draft.scheduledStart
          ? new Date(draft.scheduledStart)
          : undefined,
        scheduledEnd: draft.scheduledEnd
          ? new Date(draft.scheduledEnd)
          : undefined,
        contentType: ContentType.Alert, // Convert draft to alert
        targetLanguage: draft.targetLanguage,
        languageContent: [],
        languageGroup: draft.languageGroup,
        targetUsers: draft.targetUsers || [],
        targetGroups: [],
      });
      setEntryMode("scratch");
      setCreateStep("content");
      setAutoSaveStatus("pending");
      setErrors({});
    },
    [setEntryMode, setErrors, setNewAlert],
  );

  const handleLoadPreviousAlert = React.useCallback(
    async (sourceAlert: IAlertItem) => {
      setNewAlert((prev) => ({
        ...prev,
        title: sourceAlert.title,
        description: sourceAlert.description,
        AlertType: sourceAlert.AlertType,
        priority: sourceAlert.priority,
        isPinned: sourceAlert.isPinned,
        notificationType: sourceAlert.notificationType,
        linkUrl: sourceAlert.linkUrl || "",
        linkDescription: sourceAlert.linkDescription || "",
        targetSites: sourceAlert.targetSites || prev.targetSites,
        scheduledStart: sourceAlert.scheduledStart
          ? new Date(sourceAlert.scheduledStart)
          : undefined,
        scheduledEnd: sourceAlert.scheduledEnd
          ? new Date(sourceAlert.scheduledEnd)
          : undefined,
        contentType: ContentType.Alert,
        targetLanguage: sourceAlert.targetLanguage || prev.targetLanguage,
        languageContent: [],
        languageGroup: undefined,
        targetUsers: [],
        targetGroups: [],
      }));

      setErrors({});
      setEntryMode("scratch");
      setCreateStep("content");
      setAutoSaveStatus("pending");
    },
    [setEntryMode, setErrors, setNewAlert],
  );

  // Initialize language content when multi-language is enabled
  React.useEffect(() => {
    if (useMultiLanguage && newAlert.languageContent.length === 0) {
      // Start with English by default and generate a languageGroup ID
      setNewAlert((prev) => {
        const englishContent: ILanguageContent = {
          language: TargetLanguage.EnglishUS,
          title: prev.title,
          description: prev.description,
          linkDescription: prev.linkUrl ? prev.linkDescription : undefined,
          translationStatus: languagePolicy.workflow.enabled
            ? languagePolicy.workflow.defaultStatus
            : TranslationStatus.Approved,
        };
        // Generate a unique languageGroup ID for sharing resources (images) across language variants
        const languageGroup =
          prev.languageGroup ||
          `lg-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
        return { ...prev, languageContent: [englishContent], languageGroup };
      });
    } else if (!useMultiLanguage && newAlert.languageContent.length > 0) {
      // When switching back to single language, use the first language's content
      setNewAlert((prev) => {
        const firstLang = prev.languageContent[0];
        if (firstLang) {
          return {
            ...prev,
            title: firstLang.title,
            description: firstLang.description,
            linkDescription: firstLang.linkDescription || "",
            languageContent: [],
            languageGroup: undefined, // Clear languageGroup for single language
          };
        }
        return prev;
      });
    }
  }, [useMultiLanguage, newAlert.languageContent.length]);

  const handleTemplateSelect = React.useCallback(
    async (template: IAlertTemplate) => {
      try {
        // Get all templates from current site to include all fields
        const templateAlerts = await alertService.getTemplateAlerts(
          alertService.getCurrentSiteId(),
        );
        const fullTemplate = templateAlerts.find(
          (t: IAlertItem) => t.id === template.id,
        );

        if (fullTemplate) {
          // Use the full template data
          setNewAlert((prev) => ({
            ...prev,
            title: fullTemplate.title,
            description: fullTemplate.description,
            AlertType: fullTemplate.AlertType,
            priority: fullTemplate.priority,
            notificationType: fullTemplate.notificationType,
            isPinned: fullTemplate.isPinned,
            linkUrl: fullTemplate.linkUrl || "",
            linkDescription: fullTemplate.linkDescription || "",
            contentType: ContentType.Alert, // New alert should be Alert, not Template
          }));

          // If multi-language is enabled, load language variants of the template
          if (useMultiLanguage && fullTemplate.languageGroup) {
            try {
              // Filter the already-loaded templates for language variants
              const languageVariants = templateAlerts.filter(
                (a: IAlertItem) =>
                  a.languageGroup === fullTemplate.languageGroup,
              );

              // Convert to language content format
              const languageContent: ILanguageContent[] = languageVariants.map(
                (variant: IAlertItem) => ({
                  language: variant.targetLanguage,
                  title: variant.title,
                  description: variant.description,
                  linkDescription: variant.linkDescription || "",
                }),
              );

              if (languageContent.length > 0) {
                setNewAlert((prev) => ({
                  ...prev,
                  languageContent,
                  targetLanguage: languageContent[0].language, // Set primary language
                }));
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
          // Fallback to basic template data
          setNewAlert((prev) => ({
            ...prev,
            title: template.template.title,
            description: template.template.description,
            priority: template.template.priority,
            notificationType: template.template.notificationType,
            isPinned: template.template.isPinned,
            linkUrl: template.template.linkUrl || "",
            linkDescription: template.template.linkDescription || "",
            contentType: ContentType.Alert, // New alert should be Alert, not Template
          }));
        }
      } catch (error) {
        logger.error("CreateAlertTab", "Failed to load template", error);
        // Fallback to basic template data
        setNewAlert((prev) => ({
          ...prev,
          title: template.template.title,
          description: template.template.description,
          priority: template.template.priority,
          notificationType: template.template.notificationType,
          isPinned: template.template.isPinned,
          linkUrl: template.template.linkUrl || "",
          linkDescription: template.template.linkDescription || "",
          contentType: ContentType.Alert, // New alert should be Alert, not Template
        }));
      }

      setEntryMode("scratch");
      setCreateStep("content");
      setErrors({});
    },
    [alertService, setEntryMode, setErrors, setNewAlert, useMultiLanguage],
  );

  const validationSnapshot = React.useMemo(
    () =>
      validateAlertData(newAlert, {
        useMultiLanguage,
        languagePolicy,
        tenantDefaultLanguage: languageService.getTenantDefaultLanguage(),
        getString: getLocalizedValidationMessage,
        validateTargetSites: enableTargetSite,
      }),
    [
      enableTargetSite,
      languagePolicy,
      languageService,
      newAlert,
      useMultiLanguage,
    ],
  );

  const hasUnsavedChanges = React.useMemo(() => {
    const hasLanguageContent = newAlert.languageContent.some(
      (item) =>
        item.title.trim().length > 0 ||
        stripHtml(item.description).length > 0 ||
        (item.linkDescription || "").trim().length > 0,
    );

    return (
      newAlert.title.trim().length > 0 ||
      stripHtml(newAlert.description).length > 0 ||
      newAlert.linkUrl.trim().length > 0 ||
      newAlert.linkDescription.trim().length > 0 ||
      hasLanguageContent
    );
  }, [
    newAlert.description,
    newAlert.languageContent,
    newAlert.linkDescription,
    newAlert.linkUrl,
    newAlert.title,
  ]);

  React.useEffect(() => {
    if (!onDirtyStateChange) {
      return;
    }

    onDirtyStateChange(currentEntryMode === "scratch" && hasUnsavedChanges);
  }, [currentEntryMode, hasUnsavedChanges, onDirtyStateChange]);

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
    (step: CreateWizardStep): IFormErrors => {
      const fieldByStep: Record<CreateWizardStep, string[]> = {
        content: ["title", "description", "AlertType", "languageContent"],
        audience: ["linkUrl", "linkDescription", "targetSites"],
        publish: ["scheduledStart", "scheduledEnd"],
      };

      const stepErrors: IFormErrors = {};
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
        setErrors((prev) => ({ ...prev, ...currentStepErrors }));
        return;
      }

      setCreateStep(nextStep);
    },
    [createStep, getStepErrors, setErrors],
  );

  // Validation using shared utility
  const validateForm = React.useCallback((): boolean => {
    const validationErrors = validateAlertData(newAlert, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage: languageService.getTenantDefaultLanguage(),
      getString: getLocalizedValidationMessage,
      validateTargetSites: enableTargetSite,
    });

    setErrors(validationErrors);
    return Object.keys(validationErrors).length === 0;
  }, [
    newAlert,
    useMultiLanguage,
    languagePolicy,
    languageService,
    setErrors,
    enableTargetSite,
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

    setIsCreatingAlert(true);
    setCreationProgress([]);

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
        setCreationProgress([
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
        ]);
        notificationService.showSuccess(
          Text.format(
            strings.CreateAlertMultiLanguageSuccessStatus,
            createdVariants,
          ),
          strings.CreateAlertMultiLanguageSuccessTitle,
        );
      } else {
        setCreationProgress([
          {
            siteId: "success",
            siteName: strings.CreateAlertCreationSuccessStatus,
            hasAccess: true,
            canCreateAlerts: true,
            permissionLevel: "success",
            error: "",
          },
        ]);
        notificationService.showSuccess(
          strings.CreateAlertCreationSuccessStatus,
          strings.CreateAlertCreationSuccessTitle,
        );
      }

      // Reset form on success
      setNewAlert({
        title: "",
        description: "",
        AlertType: alertTypes.length > 0 ? alertTypes[0].name : "",
        priority: AlertPriority.Medium,
        isPinned: false,
        notificationType: NotificationType.Browser,
        linkUrl: "",
        linkDescription: "",
        targetSites: [],
        scheduledStart: undefined,
        scheduledEnd: undefined,
        contentType: ContentType.Alert,
        targetLanguage: TargetLanguage.All,
        languageContent: [],
        targetUsers: [],
        targetGroups: [],
      });
      setLastCreateAttemptFailed(false);
      setUseMultiLanguage(false);
      setCreateStep("content");
      setEntryMode("templates");
      setAutoSaveStatus("idle");
      setLastAutoSave(null);
      await loadPreviousAlerts();
    } catch (error) {
      logger.error("CreateAlertTab", "Error creating alert", error);
      setLastCreateAttemptFailed(true);
      setCreationProgress([
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
      ]);
      notificationService.showError(
        strings.CreateAlertCreationErrorStatus,
        strings.CreateAlertCreationErrorTitle,
      );
    } finally {
      setIsCreatingAlert(false);
    }
  }, [
    alertService,
    alertTypes,
    confirm,
    context.pageContext.user.displayName,
    languagePolicy,
    languageService,
    loadPreviousAlerts,
    newAlert,
    notificationService,
    setCreationProgress,
    setEntryMode,
    setIsCreatingAlert,
    setNewAlert,
    useMultiLanguage,
    validateForm,
  ]);

  const handleSaveAsDraft = React.useCallback(async () => {
    // Basic validation for drafts (less strict than full validation)
    if (!newAlert.title || newAlert.title.trim().length < 3) {
      setErrors({ title: strings.CreateAlertDraftTitleError });
      return;
    }

    // Ensure AlertType is set (required field even for drafts)
    if (
      !newAlert.AlertType ||
      !alertTypes.find((t) => t.name === newAlert.AlertType)
    ) {
      setErrors({
        AlertType: strings.CreateAlertAlertTypeError,
      });
      return;
    }

    setIsCreatingAlert(true);
    try {
      const draftData = AlertMutationService.buildDraftPayload(newAlert, {
        createdBy: context.pageContext.user.displayName,
      });
      await alertService.saveDraft(draftData);

      setCreationProgress([
        {
          siteId: "success",
          siteName: strings.CreateAlertDraftSaveSuccessStatus,
          hasAccess: true,
          canCreateAlerts: true,
          permissionLevel: "success",
          error: "",
        },
      ]);
      setLastCreateAttemptFailed(false);

      // Reload drafts and reset form after saving
      await loadDrafts();
      setTimeout(() => {
        resetForm();
        setCreationProgress([]);
      }, 2000);
    } catch (error) {
      logger.error("CreateAlertTab", "Error saving draft", error);
      setCreationProgress([
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
      ]);
      setLastCreateAttemptFailed(true);
    } finally {
      setIsCreatingAlert(false);
    }
  }, [
    newAlert,
    alertService,
    context.pageContext.user.displayName,
    setErrors,
    setIsCreatingAlert,
    setCreationProgress,
    loadDrafts,
    alertTypes,
  ]);

  // Auto-save draft functionality
  const autoSaveDraft = React.useCallback(async () => {
    // Only auto-save if there's meaningful content and AlertType is selected
    if (
      !newAlert.title ||
      newAlert.title.trim().length < 3 ||
      !newAlert.AlertType
    ) {
      return;
    }

    // Don't auto-save if currently creating/saving
    if (isCreatingAlert || isAutoSaving) {
      return;
    }

    setAutoSaveStatus("saving");
    setIsAutoSaving(true);
    try {
      const draftData = AlertMutationService.buildDraftPayload(newAlert, {
        id: autoSaveDraftId || undefined,
        createdBy: context.pageContext.user.displayName,
        titlePrefix: "[Auto-saved] ",
      });
      const savedDraft = await alertService.saveDraft(draftData);
      setAutoSaveDraftId(savedDraft.id);
      const savedAt = new Date();
      setLastAutoSave(savedAt);
      setAutoSaveStatus("saved");
      logger.debug("CreateAlertTab", "Auto-saved draft", {
        id: savedDraft.id,
        title: savedDraft.title,
      });
    } catch (error) {
      logger.warn("CreateAlertTab", "Auto-save failed", error);
      setAutoSaveStatus("error");
    } finally {
      setIsAutoSaving(false);
    }
  }, [
    newAlert,
    isCreatingAlert,
    isAutoSaving,
    autoSaveDraftId,
    alertService,
    context.pageContext.user.displayName,
    // Add other dependencies if needed, but these are the main ones for the body
  ]);

  // Debounced status updates to avoid flicker while user types
  React.useEffect(() => {
    const hasDraftableContent =
      newAlert.title.trim().length > 0 ||
      newAlert.description.trim().length > 0 ||
      newAlert.linkUrl.trim().length > 0;

    if (!hasDraftableContent) {
      setAutoSaveStatus("idle");
      return;
    }

    if (isCreatingAlert || isAutoSaving) {
      return;
    }

    if (autoSaveStatusDebounceRef.current) {
      window.clearTimeout(autoSaveStatusDebounceRef.current);
    }

    autoSaveStatusDebounceRef.current = window.setTimeout(() => {
      setAutoSaveStatus((prev) => (prev === "saving" ? prev : "pending"));
    }, 800);

    return () => {
      if (autoSaveStatusDebounceRef.current) {
        window.clearTimeout(autoSaveStatusDebounceRef.current);
      }
    };
  }, [newAlert, isCreatingAlert, isAutoSaving]);

  // Set up auto-save interval
  React.useEffect(() => {
    const intervalId = setInterval(autoSaveDraft, 30000); // Check every 30 seconds
    return () => clearInterval(intervalId);
  }, [autoSaveDraft]);

  const resetForm = React.useCallback(() => {
    setNewAlert({
      title: "",
      description: "",
      AlertType: alertTypes.length > 0 ? alertTypes[0].name : "",
      priority: AlertPriority.Medium,
      isPinned: false,
      notificationType: NotificationType.Browser,
      linkUrl: "",
      linkDescription: "",
      targetSites: [],
      scheduledStart: undefined,
      scheduledEnd: undefined,
      contentType: ContentType.Alert,
      targetLanguage: TargetLanguage.All,
      languageContent: [],
      targetUsers: [],
      targetGroups: [],
    });
    setErrors({});
    setUseMultiLanguage(false);
    setCreateStep("content");
    setLastCreateAttemptFailed(false);
    setEntryMode("templates");
    setAutoSaveDraftId(null);
    setLastAutoSave(null);
    setAutoSaveStatus("idle");
  }, [setNewAlert, alertTypes, setEntryMode, setErrors]);

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

    setCreateStep(stepOrder[createStepIndex - 1]);
  }, [canGoBack, createStepIndex, stepOrder]);

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
              <div className={styles.emptyState}>
                <p>{strings.CreateAlertPreviousAlertsEmptyMessage}</p>
              </div>
            ) : (
              previousAlerts.map((alert) => (
                <div key={alert.id} className={styles.alertCard}>
                  <div className={styles.alertCardHeader}>
                    <h4>{alert.title}</h4>
                    <small>{DateUtils.formatForDisplay(alert.createdDate)}</small>
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
                      onClick={() => void handleLoadPreviousAlert(alert)}
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
              <div className={styles.emptyState}>
                <p>{strings.CreateAlertDraftsEmptyMessage}</p>
              </div>
            ) : (
              drafts.map((draft) => (
                <div key={draft.id} className={styles.alertCard}>
                  <div className={styles.alertCardHeader}>
                    <h4>{draft.title}</h4>
                    <small>{DateUtils.formatForDisplay(draft.createdDate)}</small>
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
          setAlert={setNewAlert}
          errors={errors}
          setErrors={setErrors}
          alertTypes={alertTypes}
          alertTypeOptions={alertTypeOptions}
          priorityOptions={priorityOptions}
          notificationOptions={notificationOptions}
          contentTypeOptions={contentTypeOptions}
          languageOptions={languageOptions}
          supportedLanguages={supportedLanguages}
          useMultiLanguage={useMultiLanguage}
          setUseMultiLanguage={setUseMultiLanguage}
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
          setShowPreview={setShowPreview}
          imageFolderName={imageFolderName}
          createStep={createStep}
          onCreateStepChange={handleCreateStepChange}
          stepCompletion={stepCompletion}
          copilotAvailability={copilotAvailability}
          copilotAvailabilityMessage={copilotAvailabilityMessage}
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
            </>
          }
          afterActions={
            creationProgress.length > 0 ? (
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
            ) : undefined
          }
          applySelectedAlertTypeDefaultPriority={true}
        />
      )}
    </div>
  );
};

export default CreateAlertTab;
