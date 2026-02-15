import * as React from "react";
import {
  Delete24Regular,
  Edit24Regular,
  Globe24Regular,
  Save24Regular,
  Eye24Regular,
  Filter24Regular,
  Search24Regular,
  Calendar24Regular,
  ChevronDown24Regular,
  ChevronUp24Regular,
  Drafts24Regular,
  Send24Regular,
} from "@fluentui/react-icons";
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
  ISharePointSelectOption,
  SharePointPeoplePicker,
} from "../../UI/SharePointControls";
import SharePointRichTextEditor from "../../UI/SharePointRichTextEditor";
import SharePointDialog from "../../UI/SharePointDialog";
import MultiLanguageContentEditor from "../../UI/MultiLanguageContentEditor";
import AlertPreview from "../../UI/AlertPreview";
import SiteSelector from "../../UI/SiteSelector";
import AttachmentManager from "../../UI/AttachmentManager";
import ImageManager from "../../UI/ImageManager";
import { CopilotDraftControl } from "../../CopilotControls/CopilotDraftControl";
import { CopilotGovernanceControl } from "../../CopilotControls/CopilotGovernanceControl";
import {
  AlertPriority,
  NotificationType,
  IAlertType,
  TargetLanguage,
  ContentType,
  IPersonField,
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
import { ContentStatus } from "../../Alerts/IAlerts";
import { logger } from "../../Services/LoggerService";
import { NotificationService } from "../../Services/NotificationService";
import { SiteContextDetector } from "../../Utils/SiteContextDetector";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { IAlertItem } from "../../Alerts/IAlerts";
import { ImageStorageService } from "../../Services/ImageStorageService";
import { htmlSanitizer } from "../../Utils/HtmlSanitizer";
import { MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "../AlertSettings.module.scss";
import {
  validateAlertData,
  IFormErrors as IValidationErrors,
} from "../../Utils/AlertValidation";
import { useLanguageOptions } from "../../Hooks/useLanguageOptions";
import { usePriorityOptions } from "../../Hooks/usePriorityOptions";
import { useFluentDialogs } from "../../Hooks/useFluentDialogs";
import { StringUtils } from "../../Utils/StringUtils";
import { DateUtils } from "../../Utils/DateUtils";
import { CopilotService } from "../../Services/CopilotService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";

const STRINGS_DICTIONARY = strings as unknown as Record<string, string>;

export interface IEditingAlert extends Omit<
  IAlertItem,
  "scheduledStart" | "scheduledEnd"
> {
  scheduledStart?: Date;
  scheduledEnd?: Date;
  languageContent?: ILanguageContent[];
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
}

import { IFormErrors } from "./SharedTypes";
export type { IFormErrors };

export interface IManageAlertsTabProps {
  existingAlerts: IAlertItem[];
  setExistingAlerts: React.Dispatch<React.SetStateAction<IAlertItem[]>>;
  isLoadingAlerts: boolean;
  setIsLoadingAlerts: React.Dispatch<React.SetStateAction<boolean>>;
  selectedAlerts: string[];
  setSelectedAlerts: React.Dispatch<React.SetStateAction<string[]>>;
  editingAlert: IEditingAlert | null;
  setEditingAlert: React.Dispatch<React.SetStateAction<IEditingAlert | null>>;
  isEditingAlert: boolean;
  setIsEditingAlert: React.Dispatch<React.SetStateAction<boolean>>;
  alertTypes: IAlertType[];
  siteDetector: SiteContextDetector;
  alertService: SharePointAlertService;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  copilotEnabled?: boolean;
  setActiveTab: React.Dispatch<
    React.SetStateAction<"create" | "manage" | "types" | "settings">
  >;
}

const ManageAlertsTab: React.FC<IManageAlertsTabProps> = ({
  existingAlerts,
  setExistingAlerts,
  isLoadingAlerts,
  setIsLoadingAlerts,
  selectedAlerts,
  setSelectedAlerts,
  editingAlert,
  setEditingAlert,
  isEditingAlert,
  setIsEditingAlert,
  alertTypes,
  siteDetector,
  alertService,
  graphClient,
  context,
  copilotEnabled,
  setActiveTab,
}) => {
  const [editErrors, setEditErrors] = React.useState<IFormErrors>({});
  const [contentTypeFilter, setContentTypeFilter] = React.useState<
    "all" | ContentType
  >(ContentType.Alert);
  const [supportedLanguages, setSupportedLanguages] = React.useState<
    ISupportedLanguage[]
  >([]);
  const [useMultiLanguage, setUseMultiLanguage] = React.useState(false);
  const [tenantDefaultLanguage, setTenantDefaultLanguage] =
    React.useState<TargetLanguage>(TargetLanguage.EnglishUS);
  const [languagePolicy, setLanguagePolicy] = React.useState<ILanguagePolicy>(
    DEFAULT_LANGUAGE_POLICY,
  );
  const [showPreview, setShowPreview] = React.useState(true);
  const [supplementalAlertTypes, setSupplementalAlertTypes] = React.useState<
    IAlertType[]
  >([]);
  const [contentStatusFilter, setContentStatusFilter] = React.useState<
    "all" | ContentStatus
  >("all");

  // Enhanced filter states
  const [priorityFilter, setPriorityFilter] = React.useState<
    "all" | AlertPriority
  >("all");
  const [alertTypeFilter, setAlertTypeFilter] = React.useState<"all" | string>(
    "all",
  );
  const [statusFilter, setStatusFilter] = React.useState<"all" | string>("all");
  const [languageFilter, setLanguageFilter] = React.useState<
    "all" | TargetLanguage
  >("all");
  const [notificationFilter, setNotificationFilter] = React.useState<
    "all" | NotificationType
  >("all");
  const [dateFilter, setDateFilter] = React.useState<
    "all" | "today" | "week" | "month" | "custom"
  >("all");
  const [customDateFrom, setCustomDateFrom] = React.useState<string>("");
  const [customDateTo, setCustomDateTo] = React.useState<string>("");
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [showFilters, setShowFilters] = React.useState(false);
  const [alertsListId, setAlertsListId] = React.useState<string>("");
  const [editingAlertSiteId, setEditingAlertSiteId] =
    React.useState<string>("");
  const [editingAlertListId, setEditingAlertListId] =
    React.useState<string>("");

  React.useEffect(() => {
    const fetchListId = async () => {
      try {
        const id = await alertService.getAlertsListId();
        setAlertsListId(id);
      } catch (error) {
        logger.error("ManageAlertsTab", "Failed to get alerts list ID", error);
      }
    };
    fetchListId();
  }, [alertService]);

  React.useEffect(() => {
    if (!editingAlert) {
      setEditingAlertSiteId("");
      setEditingAlertListId("");
      return;
    }

    const parsed = alertService.parseAlertId(editingAlert.id);
    const siteId = parsed.siteId;
    setEditingAlertSiteId(siteId);

    const fetchEditingListId = async () => {
      try {
        const id = await alertService.getAlertsListId(siteId);
        setEditingAlertListId(id);
      } catch (error) {
        logger.warn(
          "ManageAlertsTab",
          "Failed to resolve alerts list ID for editing alert",
          { siteId, error },
        );
        setEditingAlertListId("");
      }
    };

    fetchEditingListId();
  }, [editingAlert?.id, alertService]);

  // Initialize services with useMemo to prevent recreation
  const languageService = React.useMemo(
    () => new LanguageAwarenessService(graphClient, context),
    [graphClient, context],
  );

  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );

  const imageStorageService = React.useMemo(
    () => new ImageStorageService(context),
    [context],
  );
  const copilotService = React.useMemo(
    () => new CopilotService(graphClient),
    [graphClient],
  );
  const { confirm, prompt, dialogs } = useFluentDialogs();

  // Priority options - using shared hook with localization
  const priorityOptions = usePriorityOptions();

  // Notification type options with detailed descriptions - memoized and localized
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

  // Alert type options (include legacy/custom values that might not exist locally)
  const alertTypeOptions: ISharePointSelectOption[] = React.useMemo(() => {
    const optionMap = new Map<string, ISharePointSelectOption>();

    [...alertTypes, ...supplementalAlertTypes].forEach((type) => {
      if (type?.name && !optionMap.has(type.name)) {
        optionMap.set(type.name, { value: type.name, label: type.name });
      }
    });

    if (editingAlert?.AlertType && !optionMap.has(editingAlert.AlertType)) {
      optionMap.set(editingAlert.AlertType, {
        value: editingAlert.AlertType,
        label: `${editingAlert.AlertType} (from source site)`,
      });
    }

    return Array.from(optionMap.values());
  }, [alertTypes, supplementalAlertTypes, editingAlert?.AlertType]);

  // Content type options - memoized and localized (matching CreateAlerts)
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
      {
        value: ContentType.Draft,
        label: strings.CreateAlertContentTypeDraftDescription,
      },
    ],
    [],
  );

  // Ensure we always have a persisted alert type selected
  React.useEffect(() => {
    if (
      editingAlert &&
      (!editingAlert.AlertType || editingAlert.AlertType.trim() === "") &&
      alertTypeOptions.length > 0
    ) {
      setEditingAlert((prev) =>
        prev ? { ...prev, AlertType: alertTypeOptions[0].value } : null,
      );
    }
  }, [editingAlert, alertTypeOptions, setEditingAlert]);

  // Language options - using shared hook
  const languageOptions = useLanguageOptions(supportedLanguages);

  // Load supported languages and tenant default on component mount
  React.useEffect(() => {
    const loadLanguageSettings = async () => {
      try {
        const baseLanguages = LanguageAwarenessService.getSupportedLanguages();
        const supportedLanguageCodes =
          await alertService.getSupportedLanguages();

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
        const defaultLang = languageService.getTenantDefaultLanguage();
        setTenantDefaultLanguage(defaultLang);
        const policy = await alertService.getLanguagePolicy();
        setLanguagePolicy(policy);
      } catch (error) {
        logger.error(
          "ManageAlertsTab",
          "Error loading language settings",
          error,
        );
        const fallbackLanguages =
          LanguageAwarenessService.getSupportedLanguages();
        setSupportedLanguages(
          fallbackLanguages.map((lang) => ({
            ...lang,
            columnExists: lang.code === TargetLanguage.EnglishUS,
            isSupported: lang.code === TargetLanguage.EnglishUS,
          })),
        );
      }
    };
    loadLanguageSettings();
  }, [languageService, alertService]);

  // Initialize language content when switching to multi-language mode in edit
  React.useEffect(() => {
    if (
      editingAlert &&
      useMultiLanguage &&
      (!editingAlert.languageContent ||
        editingAlert.languageContent.length === 0)
    ) {
      // User toggled multi-language mode - initialize with current alert content
      const initialContent: ILanguageContent = {
        language: editingAlert.targetLanguage || tenantDefaultLanguage,
        title: editingAlert.title,
        description: editingAlert.description,
        linkDescription: editingAlert.linkDescription || undefined,
        availableForAll: editingAlert.targetLanguage === tenantDefaultLanguage,
        translationStatus: languagePolicy.workflow.enabled
          ? languagePolicy.workflow.defaultStatus
          : TranslationStatus.Approved,
      };

      setEditingAlert((prev) =>
        prev
          ? {
              ...prev,
              languageContent: [initialContent],
              languageGroup:
                prev.languageGroup ||
                `lg-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`,
            }
          : null,
      );
    } else if (
      editingAlert &&
      !useMultiLanguage &&
      editingAlert.languageContent &&
      editingAlert.languageContent.length > 0
    ) {
      // User toggled back to single language - use first language content
      const firstLang = editingAlert.languageContent[0];
      if (firstLang) {
        setEditingAlert((prev) =>
          prev
            ? {
                ...prev,
                title: firstLang.title,
                description: firstLang.description,
                linkDescription: firstLang.linkDescription || "",
                targetLanguage: firstLang.language,
                languageContent: undefined,
              }
            : null,
        );
      }
    }
  }, [
    useMultiLanguage,
    editingAlert?.id,
    tenantDefaultLanguage,
    languagePolicy,
  ]); // Only run when useMultiLanguage changes or alert ID changes

  const loadExistingAlerts = React.useCallback(async () => {
    setIsLoadingAlerts(true);
    try {
      const allItems = await alertService.getAlertsAndTemplates();
      setExistingAlerts(allItems);
    } catch (error) {
      logger.error(
        "ManageAlertsTab",
        "Error loading alerts and templates",
        error,
      );
      setExistingAlerts([]);
    } finally {
      setIsLoadingAlerts(false);
    }
  }, [alertService]);

  const handleBulkDelete = React.useCallback(async () => {
    if (selectedAlerts.length === 0) return;

    const shouldDelete = await confirm({
      title: "Delete Alerts",
      message: `Are you sure you want to delete ${selectedAlerts.length} alert(s)? This action cannot be undone.`,
      confirmText: "Delete",
    });
    if (!shouldDelete) {
      return;
    }

    try {
      await Promise.all(
        selectedAlerts.map((alertId) => alertService.deleteAlert(alertId)),
      );

      // Refresh the alerts list
      await loadExistingAlerts();
      setSelectedAlerts([]);
      notificationService.showSuccess(
        `Successfully deleted ${selectedAlerts.length} alert(s)`,
        "Alerts Deleted",
      );
    } catch (error) {
      logger.error("ManageAlertsTab", "Error deleting alerts", error);
      notificationService.showError(
        "Failed to delete some alerts. Please try again.",
        "Deletion Failed",
      );
    }
  }, [selectedAlerts, alertService, loadExistingAlerts, confirm]);

  const handleEditAlert = React.useCallback(
    (alert: IAlertItem) => {
      logger.info("ManageAlertsTab", "Opening edit dialog for alert", {
        id: alert.id,
        title: alert.title,
        AlertType: alert.AlertType,
      });

      try {
        setEditErrors({});
        setSupplementalAlertTypes([]);
        const siteIdForAlert = alertService.getAlertSiteId(alert.id);
        alertService
          .getAlertTypes(siteIdForAlert)
          .then((types) => setSupplementalAlertTypes(types))
          .catch((error) => {
            logger.warn(
              "ManageAlertsTab",
              "Failed to load alert types for alert site",
              { siteIdForAlert, error },
            );
          });
        alertService
          .getLanguagePolicy(siteIdForAlert)
          .then((policy) => setLanguagePolicy(policy))
          .catch((error) => {
            logger.warn(
              "ManageAlertsTab",
              "Failed to load language policy for alert site",
              { siteIdForAlert, error },
            );
          });

        // Check if this is a multi-language alert (has languageGroup)
        const isMultiLang = !!alert.languageGroup;

        if (isMultiLang && alert.languageGroup) {
          // Load all language variants for this group
          const languageVariants = existingAlerts.filter(
            (a) => a.languageGroup === alert.languageGroup,
          );
          logger.debug("ManageAlertsTab", "Found language variants", {
            languageGroup: alert.languageGroup,
            variantCount: languageVariants.length,
          });
          const languageContent = (
            languageService?.getLanguageContent(
              languageVariants,
              alert.languageGroup,
            ) || []
          ).map((content) => ({
            ...content,
            translationStatus:
              content.translationStatus ||
              (languagePolicy.workflow.enabled
                ? languagePolicy.workflow.defaultStatus
                : TranslationStatus.Approved),
          }));

          const editingData: IEditingAlert = {
            ...alert,
            scheduledStart: alert.scheduledStart
              ? new Date(alert.scheduledStart)
              : undefined,
            scheduledEnd: alert.scheduledEnd
              ? new Date(alert.scheduledEnd)
              : undefined,
            languageContent,
            targetUsers: alert.targetUsers?.filter((u) => !u.isGroup) || [],
            targetGroups: alert.targetUsers?.filter((u) => u.isGroup) || [],
          };

          logger.info("ManageAlertsTab", "Multi-language edit mode activated", {
            languageVariants: languageVariants.length,
            editingAlertType: editingData.AlertType,
          });
          setEditingAlert(editingData);
          setUseMultiLanguage(true);

          // Duplicate language check removed
        } else {
          // Single language alert
          const editingData: IEditingAlert = {
            ...alert,
            scheduledStart: alert.scheduledStart
              ? new Date(alert.scheduledStart)
              : undefined,
            scheduledEnd: alert.scheduledEnd
              ? new Date(alert.scheduledEnd)
              : undefined,
            languageContent: undefined,
            targetUsers: alert.targetUsers?.filter((u) => !u.isGroup) || [],
            targetGroups: alert.targetUsers?.filter((u) => u.isGroup) || [],
          };

          logger.info(
            "ManageAlertsTab",
            "Single-language edit mode activated",
            {
              editingAlertType: editingData.AlertType,
            },
          );
          setEditingAlert(editingData);
          setUseMultiLanguage(false);
        }

        setShowPreview(true); // Ensure preview is visible when opening edit
      } catch (error) {
        logger.error("ManageAlertsTab", "Error opening edit dialog", error);
        notificationService.showError(
          "Failed to open edit dialog. Please try again.",
          "Edit Failed",
        );
      }
    },
    [
      setEditingAlert,
      existingAlerts,
      languageService,
      notificationService,
      alertService,
      setSupplementalAlertTypes,
      languagePolicy,
      setEditErrors,
    ],
  );

  const handlePeoplePickerChange = React.useCallback(
    (items: any[]) => {
      const users: IPersonField[] = [];
      const groups: IPersonField[] = [];

      items.forEach((item) => {
        const personField: IPersonField = {
          id: item.id,
          displayName: item.text,
          email: item.secondaryText, // Usually email
          loginName: item.loginName,
          isGroup: item.id ? item.id.indexOf("c:0(.s|true") === -1 : false, // Simple check
        };

        if (
          item.imageInitials ||
          (item.secondaryText && item.secondaryText.indexOf("@") > -1)
        ) {
          personField.isGroup = false;
          users.push(personField);
        } else {
          personField.isGroup = true;
          groups.push(personField);
        }
      });

      setEditingAlert((prev) =>
        prev
          ? {
              ...prev,
              targetUsers: users,
              targetGroups: groups,
            }
          : null,
      );
    },
    [setEditingAlert],
  );

  const handleDeleteAlert = React.useCallback(
    async (alertId: string, alertTitle: string) => {
      const shouldDelete = await confirm({
        title: "Delete Alert",
        message: `Are you sure you want to delete "${alertTitle}"? This action cannot be undone.`,
        confirmText: "Delete",
      });
      if (!shouldDelete) {
        return;
      }

      try {
        // Find the alert to get its folder name
        const alert = existingAlerts.find((a) => a.id === alertId);
        const folderName = alert?.languageGroup || alertTitle;

        // Delete the alert from SharePoint list
        await alertService.deleteAlert(alertId);

        // Delete the associated image folder
        await imageStorageService.deleteImageFolder(folderName);

        await loadExistingAlerts();
        notificationService.showSuccess(
          `Successfully deleted "${alertTitle}"`,
          "Alert Deleted",
        );
      } catch (error) {
        logger.error("ManageAlertsTab", "Error deleting alert", error);
        notificationService.showError(
          "Failed to delete alert. Please try again.",
          "Deletion Failed",
        );
      }
    },
    [
      alertService,
      loadExistingAlerts,
      notificationService,
      existingAlerts,
      context,
      confirm,
    ],
  );

  const handlePublishDraft = React.useCallback(
    async (draft: IAlertItem) => {
      const shouldPublish = await confirm({
        title: "Publish Draft",
        message: `Publish "${draft.title}" as a live alert? It will be visible to users immediately.`,
        confirmText: "Publish",
      });
      if (!shouldPublish) {
        return;
      }

      try {
        // Update the draft to convert it to an alert
        const publishedAlert: Partial<IAlertItem> = {
          ...draft,
          contentType: ContentType.Alert,
          status: "Active" as any,
        };

        await alertService.updateAlert(draft.id, publishedAlert);
        await loadExistingAlerts();
        notificationService.showSuccess(
          `Successfully published "${draft.title}"`,
          "Draft Published",
        );
      } catch (error) {
        logger.error("ManageAlertsTab", "Error publishing draft", error);
        notificationService.showError(
          "Failed to publish draft. Please try again.",
          "Publish Failed",
        );
      }
    },
    [alertService, loadExistingAlerts, notificationService, confirm],
  );

  const handleSubmitForReview = React.useCallback(
    async (alert: IAlertItem) => {
      const shouldSubmit = await confirm({
        title: "Submit for Review",
        message: `Submit "${alert.title}" for review?`,
        confirmText: "Submit",
      });
      if (!shouldSubmit) {
        return;
      }

      try {
        await alertService.submitAlert(alert.id);
        await loadExistingAlerts();
        notificationService.showSuccess(
          `Successfully submitted "${alert.title}" for review`,
          "Alert Submitted",
        );
      } catch (error) {
        logger.error("ManageAlertsTab", "Error submitting alert", error);
        notificationService.showError(
          "Failed to submit alert. Please try again.",
          "Submission Failed",
        );
      }
    },
    [alertService, loadExistingAlerts, notificationService, confirm],
  );

  const handleApprove = React.useCallback(
    async (alert: IAlertItem) => {
      const shouldApprove = await confirm({
        title: "Approve Alert",
        message: `Approve "${alert.title}"? It will become visible to users based on scheduling.`,
        confirmText: "Approve",
      });
      if (!shouldApprove) {
        return;
      }

      try {
        await alertService.approveAlert(alert.id);
        await loadExistingAlerts();
        notificationService.showSuccess(
          `Successfully approved "${alert.title}"`,
          "Alert Approved",
        );
      } catch (error) {
        logger.error("ManageAlertsTab", "Error approving alert", error);
        notificationService.showError(
          "Failed to approve alert. Please try again.",
          "Approval Failed",
        );
      }
    },
    [alertService, loadExistingAlerts, notificationService, confirm],
  );

  const handleReject = React.useCallback(
    async (alert: IAlertItem) => {
      const notes = await prompt({
        title: "Reject Alert",
        message: `Enter rejection notes for "${alert.title}" (optional):`,
        confirmText: "Reject",
        label: "Notes",
        placeholder: "Enter rejection notes...",
      });
      if (notes === null) return;

      try {
        await alertService.rejectAlert(alert.id, notes);
        await loadExistingAlerts();
        notificationService.showSuccess(
          `Rejected "${alert.title}"`,
          "Alert Rejected",
        );
      } catch (error) {
        logger.error("ManageAlertsTab", "Error rejecting alert", error);
        notificationService.showError(
          "Failed to reject alert. Please try again.",
          "Rejection Failed",
        );
      }
    },
    [alertService, loadExistingAlerts, notificationService, prompt],
  );

  // Validation using shared utility with localization
  const validateEditForm = React.useCallback((): boolean => {
    if (!editingAlert) return false;

    const validationErrors = validateAlertData(editingAlert, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage,
      getString: (key: string, ...args: Array<string | number>) => {
        const template = STRINGS_DICTIONARY[key] ?? key;
        if (args.length === 0) {
          return template;
        }
        const formattedArgs = args.map((arg) => arg.toString());
        return Text.format(template, ...formattedArgs);
      },
    });

    setEditErrors(validationErrors);
    return Object.keys(validationErrors).length === 0;
  }, [editingAlert, useMultiLanguage, languagePolicy, tenantDefaultLanguage]);

  const handleSaveEdit = React.useCallback(async () => {
    if (!editingAlert || !validateEditForm()) return;

    setIsEditingAlert(true);
    try {
      if (
        useMultiLanguage &&
        editingAlert.languageContent &&
        editingAlert.languageContent.length > 0
      ) {
        // Update multi-language alert - need to update all language variants
        if (editingAlert.languageGroup) {
          // Get all alerts in this language group and deduplicate by language
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

          // Create updated alert items
          const updatedAlerts = editingAlert.languageContent.map((content) => ({
            ...editingAlert,
            title: content.title,
            description: content.description,
            linkDescription: content.linkDescription || "",
            targetLanguage: content.language,
            availableForAll: content.availableForAll,
            translationStatus:
              content.translationStatus ||
              (languagePolicy.workflow.enabled
                ? languagePolicy.workflow.defaultStatus
                : TranslationStatus.Approved),
          }));

          // Update each language variant
          for (let i = 0; i < updatedAlerts.length; i++) {
            const updatedAlert = updatedAlerts[i];
            const existingAlert = groupAlerts.find(
              (a) => a.targetLanguage === updatedAlert.targetLanguage,
            );

            if (existingAlert) {
              // Update existing language variant
              await alertService.updateAlert(existingAlert.id, {
                title: updatedAlert.title,
                description: updatedAlert.description,
                AlertType: updatedAlert.AlertType,
                priority: updatedAlert.priority,
                isPinned: updatedAlert.isPinned,
                notificationType: updatedAlert.notificationType,
                linkUrl: updatedAlert.linkUrl,
                linkDescription: updatedAlert.linkDescription,
                scheduledStart: updatedAlert.scheduledStart?.toISOString(),
                scheduledEnd: updatedAlert.scheduledEnd?.toISOString(),
                availableForAll: updatedAlert.availableForAll,
                translationStatus: updatedAlert.translationStatus,
                targetSites: editingAlert.targetSites,
                targetUsers: [
                  ...(editingAlert.targetUsers || []),
                  ...(editingAlert.targetGroups || []),
                ],
              });
            } else {
              // Create new language variant
              await alertService.createAlert({
                title: updatedAlert.title,
                description: updatedAlert.description,
                AlertType: updatedAlert.AlertType,
                priority: updatedAlert.priority,
                isPinned: updatedAlert.isPinned,
                notificationType: updatedAlert.notificationType,
                linkUrl: updatedAlert.linkUrl,
                linkDescription: updatedAlert.linkDescription,
                targetSites: editingAlert.targetSites || [],
                scheduledStart: updatedAlert.scheduledStart?.toISOString(),
                scheduledEnd: updatedAlert.scheduledEnd?.toISOString(),
                contentType: updatedAlert.contentType,
                targetLanguage: updatedAlert.targetLanguage,
                languageGroup: updatedAlert.languageGroup,
                availableForAll: updatedAlert.availableForAll,
                translationStatus: updatedAlert.translationStatus,
                targetUsers: [
                  ...(editingAlert.targetUsers || []),
                  ...(editingAlert.targetGroups || []),
                ],
              });
            }
          }

          // Delete language variants that were removed from the editor
          const updatedLanguages = editingAlert.languageContent.map(
            (c) => c.language,
          );
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

          logger.info(
            "ManageAlertsTab",
            "Processing language variant deletions",
            {
              currentLanguages: allGroupAlerts.map((a) => a.targetLanguage),
              updatedLanguages,
              toDelete: toDelete.map((a) => ({
                id: a.id,
                language: a.targetLanguage,
              })),
            },
          );

          for (const alertToDelete of toDelete) {
            try {
              await alertService.deleteAlert(alertToDelete.id);
              logger.info(
                "ManageAlertsTab",
                `Successfully deleted language variant for ${alertToDelete.targetLanguage}`,
              );
            } catch (deleteError: any) {
              const statusCode =
                deleteError?.statusCode ||
                deleteError?.status ||
                deleteError?.response?.status ||
                deleteError?.code;
              const errorMessage = deleteError?.message || String(deleteError);

              const isNotFound =
                statusCode === 404 ||
                statusCode === "404" ||
                errorMessage.toLowerCase().includes("not found") ||
                errorMessage.toLowerCase().includes("does not exist");

              if (isNotFound) {
                // Item already deleted or doesn't exist - log and continue
                logger.warn(
                  "ManageAlertsTab",
                  `Language variant ${alertToDelete.id} already deleted or not found`,
                  {
                    language: alertToDelete.targetLanguage,
                    errorMessage,
                  },
                );
              } else {
                logger.error(
                  "ManageAlertsTab",
                  `Failed to delete language variant ${alertToDelete.id}`,
                  deleteError,
                );
                throw deleteError;
              }
            }
          }
        }
      } else {
        // Update single language alert
        await alertService.updateAlert(editingAlert.id, {
          title: editingAlert.title,
          description: editingAlert.description,
          AlertType: editingAlert.AlertType,
          priority: editingAlert.priority,
          isPinned: editingAlert.isPinned,
          notificationType: editingAlert.notificationType,
          linkUrl: editingAlert.linkUrl,
          linkDescription: editingAlert.linkDescription,
          scheduledStart: editingAlert.scheduledStart?.toISOString(),
          scheduledEnd: editingAlert.scheduledEnd?.toISOString(),
          targetSites: editingAlert.targetSites,
          targetUsers: [
            ...(editingAlert.targetUsers || []),
            ...(editingAlert.targetGroups || []),
          ],
        });
      }

      setEditingAlert(null);
      setEditErrors({});
      setUseMultiLanguage(false);
      await loadExistingAlerts();
      notificationService.showSuccess(
        "Alert updated successfully!",
        "Update Complete",
      );
    } catch (error) {
      logger.error("ManageAlertsTab", "Error updating alert", error);
      notificationService.showError(
        "Failed to update alert. Please try again.",
        "Update Failed",
      );
    } finally {
      setIsEditingAlert(false);
    }
  }, [
    editingAlert,
    validateEditForm,
    setIsEditingAlert,
    alertService,
    setEditingAlert,
    setEditErrors,
    loadExistingAlerts,
    useMultiLanguage,
    existingAlerts,
    notificationService,
    languagePolicy,
  ]);

  const handleCancelEdit = React.useCallback(() => {
    setEditingAlert(null);
    setEditErrors({});
    setUseMultiLanguage(false);
    setSupplementalAlertTypes([]);
  }, [setEditingAlert, setSupplementalAlertTypes]);

  // Helper function to get current alert type for preview (matching CreateAlerts)
  const getCurrentAlertType = React.useCallback((): IAlertType | undefined => {
    if (!editingAlert) return undefined;
    return alertTypes.find((type) => type.name === editingAlert.AlertType);
  }, [alertTypes, editingAlert]);

  const normalizeSiteId = React.useCallback((value?: string): string => {
    if (!value) return "";
    return value.toLowerCase().replace(/[{}]/g, "");
  }, []);

  const isCrossSiteAlert = React.useMemo(() => {
    if (!editingAlertSiteId) return false;
    const currentSiteId = normalizeSiteId(
      context.pageContext.site.id.toString(),
    );
    const editingSite = normalizeSiteId(editingAlertSiteId);
    return editingSite.length > 0 && editingSite !== currentSiteId;
  }, [editingAlertSiteId, context.pageContext.site.id, normalizeSiteId]);

  const editingItemId = React.useMemo(() => {
    if (!editingAlert) return undefined;
    const parsed = alertService.parseAlertId(editingAlert.id);
    const itemId = parseInt(parsed.itemId, 10);
    return Number.isNaN(itemId) ? undefined : itemId;
  }, [editingAlert?.id, alertService]);

  // Helper function to check if date matches filter
  const matchesDateFilter = React.useCallback(
    (alert: IAlertItem): boolean => {
      if (dateFilter === "all") return true;

      const alertDate = new Date(alert.createdDate || Date.now());
      const now = new Date();

      switch (dateFilter) {
        case "today":
          return alertDate.toDateString() === now.toDateString();
        case "week":
          const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          return alertDate >= weekAgo;
        case "month":
          const monthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
          return alertDate >= monthAgo;
        case "custom":
          if (!customDateFrom && !customDateTo) return true;
          const fromDate = customDateFrom
            ? new Date(customDateFrom)
            : new Date(0);
          const toDate = customDateTo ? new Date(customDateTo) : new Date();
          return alertDate >= fromDate && alertDate <= toDate;
        default:
          return true;
      }
    },
    [dateFilter, customDateFrom, customDateTo],
  );

  // Group alerts by language group with enhanced filtering
  const groupedAlerts = React.useMemo(() => {
    let filteredAlerts = [...existingAlerts];

    // Apply content type filter
    if (contentTypeFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.contentType === contentTypeFilter,
      );
    }

    // Apply priority filter
    if (priorityFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.priority === priorityFilter,
      );
    }

    // Apply alert type filter
    if (alertTypeFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.AlertType === alertTypeFilter,
      );
    }

    // Apply status filter
    if (statusFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.status.toLowerCase() === statusFilter.toLowerCase(),
      );
    }

    // Apply content status filter
    if (contentStatusFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => (a.contentStatus || ContentStatus.Draft) === contentStatusFilter,
      );
    }

    // Apply language filter
    if (languageFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.targetLanguage === languageFilter,
      );
    }

    // Apply notification type filter
    if (notificationFilter !== "all") {
      filteredAlerts = filteredAlerts.filter(
        (a) => a.notificationType === notificationFilter,
      );
    }

    // Apply date filter
    filteredAlerts = filteredAlerts.filter(matchesDateFilter);

    // Apply search term
    if (searchTerm.trim()) {
      const search = searchTerm.toLowerCase().trim();
      filteredAlerts = filteredAlerts.filter(
        (alert) =>
          alert.title?.toLowerCase().includes(search) ||
          alert.description?.toLowerCase().includes(search) ||
          alert.AlertType?.toLowerCase().includes(search) ||
          alert.createdBy?.toLowerCase().includes(search) ||
          alert.linkDescription?.toLowerCase().includes(search) ||
          alert.contentType?.toLowerCase().includes(search) ||
          alert.priority?.toLowerCase().includes(search) ||
          alert.status?.toLowerCase().includes(search),
      );
    }

    const groups: { [key: string]: IAlertItem[] } = {};
    const ungrouped: IAlertItem[] = [];

    filteredAlerts.forEach((alert) => {
      if (alert.languageGroup) {
        if (!groups[alert.languageGroup]) {
          groups[alert.languageGroup] = [];
        }
        groups[alert.languageGroup].push(alert);
      } else {
        ungrouped.push(alert);
      }
    });

    // Create combined display items: one item per language group, plus all ungrouped items
    const displayItems: (IAlertItem & {
      isLanguageGroup?: boolean;
      languageVariants?: IAlertItem[];
    })[] = [];

    // Add language groups (show primary language variant as the main item)
    Object.entries(groups).forEach(([languageGroup, variants]) => {
      // Use the first variant as the primary display item, with variants attached
      const primaryVariant =
        variants.find((v) => v.targetLanguage === TargetLanguage.EnglishUS) ||
        variants[0];
      displayItems.push({
        ...primaryVariant,
        isLanguageGroup: true,
        languageVariants: variants,
      });
    });

    // Add ungrouped items
    displayItems.push(...ungrouped);

    return displayItems;
  }, [
    existingAlerts,
    contentTypeFilter,
    priorityFilter,
    alertTypeFilter,
    statusFilter,
    languageFilter,
    notificationFilter,
    dateFilter,
    customDateFrom,
    customDateTo,
    searchTerm,
    matchesDateFilter,
    contentStatusFilter,
  ]);

  // Load alerts on mount
  React.useEffect(() => {
    loadExistingAlerts();
  }, [loadExistingAlerts]);

  return (
    <>
      {!editingAlert && (
        <div className={styles.tabContent}>
          <div className={styles.tabHeader}>
            <div>
              <h3>{strings.ManageAlerts}</h3>
              <p>
                {strings.ManageAlertsSubtitle ||
                  "View, edit, and manage existing alerts across your sites"}
              </p>
            </div>
            <div className={styles.flexRowGap12}>
              {selectedAlerts.length === 1 && (
                <SharePointButton
                  variant="primary"
                  icon={<Edit24Regular />}
                  onClick={() => {
                    const selectedAlert = groupedAlerts.find((alert) =>
                      selectedAlerts.includes(alert.id),
                    );
                    if (selectedAlert) {
                      handleEditAlert(selectedAlert);
                    }
                  }}
                >
                  {strings.ManageAlertsEditSelectedButtonLabel ||
                    "Edit Selected"}
                </SharePointButton>
              )}
              {selectedAlerts.length > 0 && (
                <SharePointButton
                  variant="danger"
                  icon={<Delete24Regular />}
                  onClick={handleBulkDelete}
                >
                  {Text.format(
                    strings.ManageAlertsDeleteSelectedButtonLabel ||
                      "Delete Selected ({0})",
                    selectedAlerts.length,
                  )}
                </SharePointButton>
              )}
              <SharePointButton
                variant="secondary"
                onClick={loadExistingAlerts}
                disabled={isLoadingAlerts}
              >
                {isLoadingAlerts
                  ? strings.ManageAlertsRefreshingLabel || "Refreshing..."
                  : strings.Refresh}
              </SharePointButton>
              <button
                className={styles.filterToggleButton}
                onClick={() => setShowFilters(!showFilters)}
              >
                <Filter24Regular />
                {showFilters
                  ? strings.ManageAlertsHideFiltersLabel || "Hide"
                  : strings.ManageAlertsShowFiltersLabel || "Show"}{" "}
                {strings.ManageAlertsFiltersLabel || "Filters"}
                {showFilters ? (
                  <ChevronUp24Regular />
                ) : (
                  <ChevronDown24Regular />
                )}
              </button>
            </div>
          </div>

          {isLoadingAlerts ? (
            <div className={styles.alertsList}>
              {[1, 2, 3].map((i) => (
                <div key={i} className={styles.alertCard} aria-hidden="true">
                  <div className={styles.alertCardHeader}>
                    <div
                      className={styles.skeletonLine}
                      style={{ width: "60px", height: "14px" }}
                    />
                    <div
                      className={styles.skeletonLine}
                      style={{
                        width: "80px",
                        height: "20px",
                        borderRadius: "12px",
                      }}
                    />
                  </div>
                  <div className={styles.alertCardContent}>
                    <div
                      className={styles.skeletonLine}
                      style={{
                        width: "70%",
                        height: "16px",
                        marginBottom: "8px",
                      }}
                    />
                    <div
                      className={styles.skeletonLine}
                      style={{
                        width: "100%",
                        height: "14px",
                        marginBottom: "6px",
                      }}
                    />
                    <div
                      className={styles.skeletonLine}
                      style={{
                        width: "85%",
                        height: "14px",
                        marginBottom: "12px",
                      }}
                    />
                    <div className={styles.alertMetaData}>
                      <div
                        className={styles.skeletonLine}
                        style={{ width: "45%", height: "12px" }}
                      />
                      <div
                        className={styles.skeletonLine}
                        style={{ width: "50%", height: "12px" }}
                      />
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ) : existingAlerts.length === 0 ? (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>📢</div>
              <h4>{strings.ManageAlertsEmptyTitle || "No Alerts Found"}</h4>
              <p>
                {strings.ManageAlertsEmptyDescription ||
                  "No alerts are currently available. This might be because:"}
              </p>
              <ul className={styles.emptyStateList}>
                <li>
                  {strings.ManageAlertsEmptyReasonListsMissing ||
                    "The Alert Banner lists haven't been created yet"}
                </li>
                <li>
                  {strings.ManageAlertsEmptyReasonNoAccess ||
                    "You don't have access to the lists"}
                </li>
                <li>
                  {strings.ManageAlertsEmptyReasonNoAlerts ||
                    "No alerts have been created yet"}
                </li>
              </ul>
              <div className={styles.flexRowCentered}>
                <SharePointButton
                  variant="primary"
                  onClick={() => setActiveTab("create")}
                >
                  {strings.ManageAlertsEmptyCreateFirstButton ||
                    "Create First Alert"}
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={loadExistingAlerts}
                >
                  {strings.Refresh}
                </SharePointButton>
              </div>
            </div>
          ) : (
            <div className={styles.alertsList}>
              {/* Enhanced Filter Section */}
              <div className={styles.filterSection}>
                <div className={styles.filterHeader}>
                  <div className={styles.filterSummary}>
                    <span>
                      {Text.format(
                        strings.ManageAlertsFilterSummaryCount ||
                          "Showing {0} of {1} items",
                        groupedAlerts.length,
                        existingAlerts.length,
                      )}
                    </span>
                    <span>
                      {Text.format(
                        strings.ManageAlertsFilterSummaryLanguageGroups ||
                          "{0} multi-language groups",
                        groupedAlerts.filter((a) => a.isLanguageGroup).length,
                      )}
                    </span>
                  </div>
                </div>

                {/* Quick Content Type Tabs */}
                <div
                  className={styles.templateActions}
                  style={{ marginBottom: "12px" }}
                >
                  <SharePointButton
                    variant={
                      contentTypeFilter === ContentType.Alert
                        ? "primary"
                        : "secondary"
                    }
                    onClick={() => setContentTypeFilter(ContentType.Alert)}
                  >
                    {Text.format(
                      strings.ManageAlertsFilterAlertsLabel ||
                        "📢 Alerts ({0})",
                      existingAlerts.filter(
                        (a) => a.contentType === ContentType.Alert,
                      ).length,
                    )}
                  </SharePointButton>
                  <SharePointButton
                    variant={
                      contentTypeFilter === ContentType.Draft
                        ? "primary"
                        : "secondary"
                    }
                    onClick={() => setContentTypeFilter(ContentType.Draft)}
                  >
                    {Text.format(
                      strings.ManageAlertsFilterDraftsLabel ||
                        "✏️ Drafts ({0})",
                      existingAlerts.filter(
                        (a) => a.contentType === ContentType.Draft,
                      ).length,
                    )}
                  </SharePointButton>
                  <SharePointButton
                    variant={
                      contentTypeFilter === ContentType.Template
                        ? "primary"
                        : "secondary"
                    }
                    onClick={() => setContentTypeFilter(ContentType.Template)}
                  >
                    {Text.format(
                      strings.ManageAlertsFilterTemplatesLabel ||
                        "📄 Templates ({0})",
                      existingAlerts.filter(
                        (a) => a.contentType === ContentType.Template,
                      ).length,
                    )}
                  </SharePointButton>
                  <SharePointButton
                    variant={
                      contentTypeFilter === "all" ? "primary" : "secondary"
                    }
                    onClick={() => setContentTypeFilter("all")}
                  >
                    {Text.format(
                      strings.ManageAlertsFilterAllLabel || "📋 All ({0})",
                      existingAlerts.length,
                    )}
                  </SharePointButton>
                </div>

                {/* Search Bar - Always Visible */}
                <div className={styles.searchBar}>
                  <SharePointInput
                    label=""
                    value={searchTerm}
                    onChange={setSearchTerm}
                    placeholder={
                      strings.ManageAlertsSearchPlaceholder ||
                      "🔍 Search alerts, drafts, and templates by title, description, type, priority, or author..."
                    }
                    className={styles.searchInput}
                  />
                </div>

                {/* Collapsible Advanced Filters */}
                {showFilters && (
                  <div className={styles.advancedFilters}>
                    <div className={styles.filterGrid}>
                      {/* Content Type Filter */}
                      <SharePointSelect
                        label={strings.ContentTypeLabel}
                        value={contentTypeFilter}
                        onChange={(value) =>
                          setContentTypeFilter(value as "all" | ContentType)
                        }
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsContentTypeAllOption ||
                              "📋 All Types",
                          },
                          {
                            value: ContentType.Alert,
                            label:
                              strings.ManageAlertsContentTypeAlertOption ||
                              "📢 Alerts",
                          },
                          {
                            value: ContentType.Draft,
                            label:
                              strings.ManageAlertsContentTypeDraftOption ||
                              "✏️ Drafts",
                          },
                          {
                            value: ContentType.Template,
                            label:
                              strings.ManageAlertsContentTypeTemplateOption ||
                              "📄 Templates",
                          },
                        ]}
                      />

                      {/* Priority Filter */}
                      <SharePointSelect
                        label={strings.CreateAlertPriorityLabel}
                        value={priorityFilter}
                        onChange={(value) =>
                          setPriorityFilter(value as "all" | AlertPriority)
                        }
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsPriorityAllOption ||
                              "🔘 All Priorities",
                          },
                          {
                            value: AlertPriority.Critical,
                            label: `🔴 ${strings.PriorityCritical}`,
                          },
                          {
                            value: AlertPriority.High,
                            label: `🟠 ${strings.PriorityHigh}`,
                          },
                          {
                            value: AlertPriority.Medium,
                            label: `🟡 ${strings.PriorityMedium}`,
                          },
                          {
                            value: AlertPriority.Low,
                            label: `🟢 ${strings.PriorityLow}`,
                          },
                        ]}
                      />

                      {/* Alert Type Filter */}
                      <SharePointSelect
                        label={strings.AlertType}
                        value={alertTypeFilter}
                        onChange={(value) => setAlertTypeFilter(value)}
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsAlertTypeAllOption ||
                              "🎨 All Types",
                          },
                          ...alertTypes.map((type) => ({
                            value: type.name,
                            label: `${type.iconName ? "🎯" : "📢"} ${type.name}`,
                          })),
                        ]}
                      />

                      {/* Status Filter */}
                      <SharePointSelect
                        label={strings.Status}
                        value={statusFilter}
                        onChange={(value) => setStatusFilter(value)}
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsStatusAllOption ||
                              "⚪ All Statuses",
                          },
                          {
                            value: "active",
                            label: `🟢 ${strings.StatusActive}`,
                          },
                          {
                            value: "expired",
                            label: `🔴 ${strings.StatusExpired}`,
                          },
                          {
                            value: "scheduled",
                            label: `🟡 ${strings.StatusScheduled}`,
                          },
                        ]}
                      />

                      {/* Content Status Filter */}
                      <SharePointSelect
                        label="Approval Status"
                        value={contentStatusFilter}
                        onChange={(value) =>
                          setContentStatusFilter(value as ContentStatus | "all")
                        }
                        options={[
                          { value: "all", label: "Any Status" },
                          { value: ContentStatus.Draft, label: "✏️ Draft" },
                          {
                            value: ContentStatus.PendingReview,
                            label: "⏳ Pending Review",
                          },
                          {
                            value: ContentStatus.Approved,
                            label: "✅ Approved",
                          },
                          {
                            value: ContentStatus.Rejected,
                            label: "❌ Rejected",
                          },
                        ]}
                      />

                      {/* Language Filter */}
                      <SharePointSelect
                        label={strings.Language}
                        value={languageFilter}
                        onChange={(value) =>
                          setLanguageFilter(value as "all" | TargetLanguage)
                        }
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsLanguageAllOption ||
                              "🌐 All Languages",
                          },
                          {
                            value: TargetLanguage.All,
                            label:
                              strings.ManageAlertsLanguageMultiOption ||
                              "🌍 Multi-Language",
                          },
                          ...supportedLanguages.map((lang) => ({
                            value: lang.code,
                            label: `${lang.flag} ${lang.nativeName}`,
                          })),
                        ]}
                      />

                      {/* Notification Type Filter */}
                      <SharePointSelect
                        label={strings.NotificationType}
                        value={notificationFilter}
                        onChange={(value) =>
                          setNotificationFilter(
                            value as "all" | NotificationType,
                          )
                        }
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsNotificationAllOption ||
                              "📧 All Notification Types",
                          },
                          {
                            value: NotificationType.None,
                            label:
                              strings.ManageAlertsNotificationNoneOption ||
                              "🚫 None - Banner only",
                          },
                          {
                            value: NotificationType.Browser,
                            label:
                              strings.ManageAlertsNotificationBrowserOption ||
                              "🌐 Browser - Banner display",
                          },
                          {
                            value: NotificationType.Email,
                            label:
                              strings.ManageAlertsNotificationEmailOption ||
                              "📧 Email only - No banner",
                          },
                          {
                            value: NotificationType.Both,
                            label:
                              strings.ManageAlertsNotificationBothOption ||
                              "📧🌐 Browser + Email",
                          },
                        ]}
                      />

                      {/* Date Filter */}
                      <SharePointSelect
                        label={
                          strings.ManageAlertsCreatedDateFilterLabel ||
                          "Created Date"
                        }
                        value={dateFilter}
                        onChange={(value) => setDateFilter(value as any)}
                        options={[
                          {
                            value: "all",
                            label:
                              strings.ManageAlertsDateAllOption ||
                              "📅 All Dates",
                          },
                          {
                            value: "today",
                            label:
                              strings.ManageAlertsDateTodayOption || "📅 Today",
                          },
                          {
                            value: "week",
                            label:
                              strings.ManageAlertsDateWeekOption ||
                              "📅 This Week",
                          },
                          {
                            value: "month",
                            label:
                              strings.ManageAlertsDateMonthOption ||
                              "📅 This Month",
                          },
                          {
                            value: "custom",
                            label:
                              strings.ManageAlertsDateCustomOption ||
                              "📅 Custom Range",
                          },
                        ]}
                      />
                    </div>

                    {/* Custom Date Range */}
                    {dateFilter === "custom" && (
                      <div className={styles.dateRangeFilters}>
                        <SharePointInput
                          label={
                            strings.ManageAlertsFromDateLabel || "From Date"
                          }
                          type="date"
                          value={customDateFrom}
                          onChange={setCustomDateFrom}
                        />
                        <SharePointInput
                          label={strings.ManageAlertsToDateLabel || "To Date"}
                          type="date"
                          value={customDateTo}
                          onChange={setCustomDateTo}
                        />
                      </div>
                    )}

                    {/* Clear Filters Button */}
                    <div className={styles.filterActions}>
                      <SharePointButton
                        variant="secondary"
                        onClick={() => {
                          setContentTypeFilter(ContentType.Alert);
                          setPriorityFilter("all");
                          setAlertTypeFilter("all");
                          setStatusFilter("all");
                          setLanguageFilter("all");
                          setNotificationFilter("all");
                          setDateFilter("all");
                          setCustomDateFrom("");
                          setCustomDateTo("");
                          setSearchTerm("");
                          setContentStatusFilter("all");
                        }}
                      >
                        {strings.ManageAlertsClearFiltersButton ||
                          "Clear All Filters"}
                      </SharePointButton>
                    </div>
                  </div>
                )}
              </div>

              {groupedAlerts.map((alert) => {
                const alertType = alertTypes.find(
                  (type) => type.name === alert.AlertType,
                );
                const isSelected = selectedAlerts.includes(alert.id);
                const isMultiLanguage =
                  alert.isLanguageGroup &&
                  alert.languageVariants &&
                  alert.languageVariants.length > 1;

                // Template rendering validation
                if (
                  alert.contentType === ContentType.Template &&
                  (!alert.title || !alert.description)
                ) {
                  logger.warn(
                    "ManageAlertsTab",
                    `Template ${alert.id} has missing content`,
                    {
                      hasTitle: !!alert.title,
                      hasDescription: !!alert.description,
                    },
                  );
                }

                return (
                  <div
                    key={alert.id}
                    className={`${styles.alertCard} ${isSelected ? styles.selected : ""} ${alert.contentType === ContentType.Template ? styles.templateCard : ""}`}
                  >
                    <div className={styles.alertCardHeader}>
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedAlerts((prev) => [...prev, alert.id]);
                          } else {
                            setSelectedAlerts((prev) =>
                              prev.filter((id) => id !== alert.id),
                            );
                          }
                        }}
                        className={styles.alertCheckbox}
                      />
                      <div className={styles.alertStatus}>
                        <span
                          className={`${styles.statusBadge} ${alert.status.toLowerCase() === "active" ? styles.active : alert.status.toLowerCase() === "expired" ? styles.expired : styles.scheduled}`}
                        >
                          {alert.status}
                        </span>
                        {alert.isPinned && (
                          <span className={styles.pinnedBadge}>📌 PINNED</span>
                        )}
                        {alert.contentStatus &&
                          alert.contentStatus !== ContentStatus.Approved && (
                            <span
                              className={`${styles.statusBadge} ${styles.template}`}
                            >
                              {alert.contentStatus ===
                              ContentStatus.PendingReview
                                ? "⏳ Pending"
                                : alert.contentStatus === ContentStatus.Rejected
                                  ? "❌ Rejected"
                                  : "✏️ Draft"}
                            </span>
                          )}
                      </div>
                    </div>

                    {/* Force render alertCardContent with error boundaries for templates */}
                    <div
                      className={styles.alertCardContent}
                      style={
                        alert.contentType === ContentType.Template
                          ? {
                              display: "block",
                              visibility: "visible",
                              opacity: 1,
                            }
                          : {}
                      }
                    >
                      {/* Show AlertType if available, otherwise show content type */}
                      {alert.AlertType ? (
                        <div
                          className={styles.alertTypeIndicator}
                          style={
                            {
                              "--bg-color":
                                alertType?.backgroundColor || "#0078d4",
                              "--text-color": alertType?.textColor || "#ffffff",
                            } as React.CSSProperties
                          }
                        >
                          {alert.AlertType}
                        </div>
                      ) : (
                        <div
                          className={styles.alertTypeIndicator}
                          style={
                            {
                              "--bg-color":
                                alert.contentType === ContentType.Template
                                  ? "#8764b8"
                                  : "#0078d4",
                              "--text-color": "#ffffff",
                            } as React.CSSProperties
                          }
                        >
                          {alert.contentType === ContentType.Template
                            ? strings.ManageAlertsTemplateBadge || "📄 Template"
                            : strings.ManageAlertsAlertBadge || "📢 Alert"}
                        </div>
                      )}

                      <h4 className={styles.alertCardTitle}>
                        {alert.title ||
                          strings.ManageAlertsNoTitleFallback ||
                          "[No Title Available]"}
                        {isMultiLanguage && (
                          <span className={styles.multiLanguageBadge}>
                            <Globe24Regular
                              style={{
                                width: "12px",
                                height: "12px",
                                marginRight: "4px",
                              }}
                            />
                            {Text.format(
                              strings.ManageAlertsLanguageBadge ||
                                "{0} languages",
                              alert.languageVariants?.length || 0,
                            )}
                          </span>
                        )}
                      </h4>

                      <div className={styles.alertCardDescription}>
                        {alert.description ? (
                          <div
                            dangerouslySetInnerHTML={{
                              __html: htmlSanitizer.sanitizeHtml(
                                StringUtils.truncate(alert.description, 150),
                              ),
                            }}
                          />
                        ) : (
                          <em style={{ color: "#999" }}>
                            {strings.ManageAlertsNoDescriptionFallback ||
                              "[No Description Available]"}
                          </em>
                        )}
                      </div>

                      <div className={styles.alertMetaData}>
                        <div className={styles.metaInfo}>
                          <strong>{strings.AlertType || "Type"}:</strong>
                          <span
                            className={`${styles.contentTypeBadge} ${alert.contentType === ContentType.Template ? styles.template : styles.alert}`}
                          >
                            {alert.contentType === ContentType.Template
                              ? strings.ManageAlertsTemplateBadge ||
                                "📄 Template"
                              : strings.ManageAlertsAlertBadge || "📢 Alert"}
                          </span>
                        </div>
                        <div className={styles.metaInfo}>
                          <strong>{strings.Priority || "Priority"}:</strong>{" "}
                          {priorityOptions.find(
                            (option) => option.value === alert.priority,
                          )?.label || alert.priority}
                        </div>
                        <div className={styles.metaInfo}>
                          <strong>{strings.Language || "Language"}:</strong>
                          {isMultiLanguage ? (
                            <span className={styles.languageList}>
                              {Text.format(
                                strings.ManageAlertsMultiLanguageList ||
                                  "🌐 Multi-language ({0})",
                                alert.languageVariants
                                  ?.map(
                                    (v) =>
                                      supportedLanguages.find(
                                        (l) => l.code === v.targetLanguage,
                                      )?.flag || v.targetLanguage,
                                  )
                                  .join(" ") || "",
                              )}
                            </span>
                          ) : alert.targetLanguage === TargetLanguage.All ? (
                            strings.ManageAlertsAllLanguagesLabel ||
                            "🌐 All Languages"
                          ) : (
                            `${supportedLanguages.find((l) => l.code === alert.targetLanguage)?.flag || ""} ${supportedLanguages.find((l) => l.code === alert.targetLanguage)?.nativeName || alert.targetLanguage}`
                          )}
                        </div>
                        {alert.linkUrl && (
                          <div className={styles.metaInfo}>
                            <strong>
                              {strings.LinkDescription || "Action"}:
                            </strong>{" "}
                            {alert.linkDescription}
                          </div>
                        )}
                        <div className={styles.metaInfo}>
                          <strong>
                            {strings.ManageAlertsCreatedLabel || "Created"}:
                          </strong>{" "}
                          {DateUtils.formatForDisplay(
                            alert.createdDate || new Date(),
                          )}
                        </div>
                        {alert.scheduledStart && (
                          <div className={styles.metaInfo}>
                            <strong>
                              {strings.CreateAlertStartDateLabel || "Start"}:
                            </strong>{" "}
                            {DateUtils.formatDateTimeForDisplay(
                              alert.scheduledStart,
                            )}
                          </div>
                        )}
                        {alert.scheduledEnd && (
                          <div className={styles.metaInfo}>
                            <strong>
                              {strings.CreateAlertEndDateLabel || "End"}:
                            </strong>{" "}
                            {DateUtils.formatDateTimeForDisplay(
                              alert.scheduledEnd,
                            )}
                          </div>
                        )}
                      </div>
                    </div>

                    <div className={styles.alertCardActions}>
                      {/* Approval Workflow Buttons */}
                      {(alert.contentStatus === ContentStatus.Draft ||
                        alert.contentStatus === ContentStatus.Rejected ||
                        !alert.contentStatus) && (
                        <SharePointButton
                          variant="primary"
                          title="Submit for Review"
                          onClick={() => handleSubmitForReview(alert)}
                        >
                          Submit
                        </SharePointButton>
                      )}
                      {alert.contentStatus === ContentStatus.PendingReview && (
                        <>
                          <SharePointButton
                            variant="primary"
                            onClick={() => handleApprove(alert)}
                          >
                            Approve
                          </SharePointButton>
                          <SharePointButton
                            variant="danger"
                            onClick={() => handleReject(alert)}
                          >
                            Reject
                          </SharePointButton>
                        </>
                      )}

                      <SharePointButton
                        variant="secondary"
                        icon={<Edit24Regular />}
                        onClick={() => {
                          handleEditAlert(alert);
                        }}
                      >
                        {strings.Edit}
                      </SharePointButton>
                      <SharePointButton
                        variant="danger"
                        icon={<Delete24Regular />}
                        onClick={() => {
                          handleDeleteAlert(alert.id, alert.title);
                        }}
                      >
                        {strings.Delete}
                      </SharePointButton>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* Edit Alert Form — rendered inline, replaces the parent dialog content */}
      {editingAlert && (
        <div className={styles.alertForm}>
          <div className={styles.formWithPreview}>
            <div className={styles.formColumn}>
              <SharePointSection
                title={strings.CreateAlertSectionContentClassificationTitle}
              >
                <SharePointSelect
                  label={strings.ContentTypeLabel}
                  value={editingAlert.contentType}
                  onChange={(value) =>
                    setEditingAlert((prev) =>
                      prev
                        ? { ...prev, contentType: value as ContentType }
                        : null,
                    )
                  }
                  options={contentTypeOptions}
                  required
                  description={
                    strings.ManageAlertsContentClassificationDescription ||
                    strings.CreateAlertSectionContentClassificationDescription
                  }
                />

                <div className={styles.languageModeSelector}>
                  <label className={styles.fieldLabel}>
                    {strings.ManageAlertsLanguageConfigurationLabel ||
                      strings.CreateAlertLanguageConfigurationLabel}
                  </label>
                  <div className={styles.languageOptions}>
                    <SharePointButton
                      variant={!useMultiLanguage ? "primary" : "secondary"}
                      onClick={() => setUseMultiLanguage(false)}
                    >
                      {strings.ManageAlertsSingleLanguageButton ||
                        "🌐 Single Language"}
                    </SharePointButton>
                    <SharePointButton
                      variant={useMultiLanguage ? "primary" : "secondary"}
                      onClick={() => setUseMultiLanguage(true)}
                    >
                      {strings.ManageAlertsMultiLanguageButton ||
                        "🗣️ Multi-Language"}
                    </SharePointButton>
                  </div>
                </div>
              </SharePointSection>
              <SharePointSection
                title={strings.CreateAlertSectionUserTargetingTitle}
              >
                <SharePointPeoplePicker
                  context={context}
                  titleText={strings.CreateAlertPeoplePickerLabel}
                  personSelectionLimit={50}
                  groupName={strings.CreateAlertPeoplePickerLabel}
                  showtooltip={true}
                  defaultSelectedUsers={[
                    ...(editingAlert.targetUsers || []),
                    ...(editingAlert.targetGroups || []),
                  ].map((u) => u.email || u.loginName || u.displayName)}
                  onChange={handlePeoplePickerChange}
                  principalTypes={[
                    PrincipalType.User,
                    PrincipalType.SharePointGroup,
                    PrincipalType.SecurityGroup,
                    PrincipalType.DistributionList,
                  ]}
                  resolveDelay={1000}
                  description={strings.CreateAlertPeoplePickerDescription}
                />
              </SharePointSection>
              {!useMultiLanguage ? (
                <>
                  <SharePointSection
                    title={strings.CreateAlertSectionLanguageTargetingTitle}
                  >
                    <SharePointSelect
                      label={strings.CreateAlertTargetLanguageLabel}
                      value={editingAlert.targetLanguage}
                      onChange={(value) =>
                        setEditingAlert((prev) =>
                          prev
                            ? {
                                ...prev,
                                targetLanguage: value as TargetLanguage,
                              }
                            : null,
                        )
                      }
                      options={languageOptions}
                      required
                      description={
                        strings.ManageAlertsTargetLanguageDescription ||
                        strings.CreateAlertSectionLanguageTargetingDescription
                      }
                    />
                  </SharePointSection>

                  <SharePointSection
                    title={strings.CreateAlertSectionBasicInformationTitle}
                  >
                    <SharePointInput
                      label={strings.AlertTitle}
                      value={editingAlert.title}
                      onChange={(value) => {
                        setEditingAlert((prev) =>
                          prev ? { ...prev, title: value } : null,
                        );
                        if (editErrors.title)
                          setEditErrors((prev) => ({
                            ...prev,
                            title: undefined,
                          }));
                      }}
                      placeholder={strings.CreateAlertTitlePlaceholder}
                      required
                      error={editErrors.title}
                      description={strings.CreateAlertTitleDescription}
                    />
                    {copilotEnabled && (
                      <div className={styles.copilotActionsRow}>
                        <CopilotDraftControl
                          copilotService={copilotService}
                          onDraftGenerated={(draft) =>
                            setEditingAlert((prev) =>
                              prev ? { ...prev, description: draft } : null,
                            )
                          }
                          onError={(error) =>
                            notificationService.showError(
                              error,
                              strings.CopilotErrorTitle || "Copilot Error",
                            )
                          }
                          disabled={!editingAlert || isEditingAlert}
                        />
                      </div>
                    )}

                    <SharePointRichTextEditor
                      label={strings.AlertDescription}
                      value={editingAlert.description}
                      onChange={(value) => {
                        setEditingAlert((prev) =>
                          prev ? { ...prev, description: value } : null,
                        );
                        if (editErrors.description)
                          setEditErrors((prev) => ({
                            ...prev,
                            description: undefined,
                          }));
                      }}
                      context={context}
                      placeholder={
                        strings.ManageAlertsDescriptionPlaceholder ||
                        strings.CreateAlertDescriptionPlaceholder
                      }
                      required
                      error={editErrors.description}
                      description={
                        strings.ManageAlertsDescriptionHelp ||
                        strings.CreateAlertDescriptionHelp
                      }
                      imageFolderName={
                        editingAlert.languageGroup ||
                        editingAlert.title ||
                        "Untitled_Alert"
                      }
                      disableImageUpload={isCrossSiteAlert}
                    />
                    {copilotEnabled && (
                      <CopilotGovernanceControl
                        copilotService={copilotService}
                        textToAnalyze={editingAlert.description}
                        onError={(error) =>
                          notificationService.showError(
                            error,
                            strings.CopilotErrorTitle || "Copilot Error",
                          )
                        }
                        disabled={!editingAlert || isEditingAlert}
                      />
                    )}
                  </SharePointSection>
                </>
              ) : (
                <SharePointSection title={strings.MultiLanguageContent}>
                  <MultiLanguageContentEditor
                    content={editingAlert.languageContent || []}
                    onContentChange={(content) => {
                      setEditingAlert((prev) =>
                        prev ? { ...prev, languageContent: content } : null,
                      );
                    }}
                    availableLanguages={supportedLanguages}
                    errors={editErrors}
                    linkUrl={editingAlert.linkUrl}
                    tenantDefaultLanguage={tenantDefaultLanguage}
                    languagePolicy={languagePolicy}
                    context={context}
                    imageFolderName={editingAlert.languageGroup}
                    disableImageUpload={isCrossSiteAlert}
                    copilotService={copilotEnabled ? copilotService : undefined}
                  />
                </SharePointSection>
              )}
              {/* Image Manager */}
              {editingAlert.languageGroup && !isCrossSiteAlert && (
                <SharePointSection
                  title={strings.ImageManagerTitle || "Images"}
                >
                  <ImageManager
                    context={context}
                    imageStorageService={imageStorageService}
                    folderName={editingAlert.languageGroup}
                  />
                </SharePointSection>
              )}
              {editingAlert.languageGroup && isCrossSiteAlert && (
                <SharePointSection
                  title={strings.ImageManagerTitle || "Images"}
                >
                  <div className={styles.infoMessage}>
                    Image management is disabled for alerts created on other
                    sites. Open the source site to manage images.
                  </div>
                </SharePointSection>
              )}
              // ... (rest of the component)
              {/* Attachment Manager */}
              <AttachmentManager
                context={context}
                alertService={alertService}
                listId={editingAlertListId || alertsListId}
                itemId={editingItemId}
                siteId={editingAlertSiteId || undefined}
                attachments={editingAlert.attachments || []}
                onAttachmentsChange={(newAttachments) => {
                  setEditingAlert((prev) =>
                    prev ? { ...prev, attachments: newAttachments } : null,
                  );
                }}
              />
              {/* Image Manager */}
              <SharePointSection
                title={
                  strings.ManageAlertsUploadedImagesSectionTitle ||
                  "Uploaded Images"
                }
              >
                {!isCrossSiteAlert ? (
                  <ImageManager
                    context={context}
                    imageStorageService={imageStorageService}
                    folderName={
                      editingAlert.languageGroup || editingAlert.title
                    }
                    onImageDeleted={() => {
                      // Optional: Refresh or notify on image deletion
                      logger.info(
                        "ManageAlertsTab",
                        "Image deleted from alert folder",
                      );
                    }}
                  />
                ) : (
                  <div className={styles.infoMessage}>
                    Image management is disabled for alerts created on other
                    sites. Open the source site to manage images.
                  </div>
                )}
              </SharePointSection>
              <SharePointSection
                title={strings.CreateAlertConfigurationSectionTitle}
              >
                <SharePointSelect
                  label={strings.AlertType}
                  value={editingAlert.AlertType}
                  onChange={(value) => {
                    setEditingAlert((prev) =>
                      prev ? { ...prev, AlertType: value } : null,
                    );
                    if (editErrors.AlertType)
                      setEditErrors((prev) => ({
                        ...prev,
                        AlertType: undefined,
                      }));
                  }}
                  options={alertTypeOptions}
                  required
                  error={editErrors.AlertType}
                  description={
                    strings.ManageAlertsAlertTypeDescription ||
                    strings.CreateAlertConfigurationDescription
                  }
                />

                <SharePointSelect
                  label={strings.CreateAlertPriorityLabel}
                  value={editingAlert.priority}
                  onChange={(value) =>
                    setEditingAlert((prev) =>
                      prev
                        ? { ...prev, priority: value as AlertPriority }
                        : null,
                    )
                  }
                  options={priorityOptions}
                  required
                  description={
                    strings.ManageAlertsPriorityDescription ||
                    strings.CreateAlertPriorityDescription
                  }
                />

                <SharePointToggle
                  label={strings.CreateAlertPinLabel}
                  checked={editingAlert.isPinned}
                  onChange={(checked) =>
                    setEditingAlert((prev) =>
                      prev ? { ...prev, isPinned: checked } : null,
                    )
                  }
                  description={
                    strings.ManageAlertsPinDescription ||
                    strings.CreateAlertPinDescription
                  }
                />

                <SharePointSelect
                  label={
                    strings.ManageAlertsNotificationMethodLabel ||
                    strings.CreateAlertNotificationLabel
                  }
                  value={editingAlert.notificationType}
                  onChange={(value) =>
                    setEditingAlert((prev) =>
                      prev
                        ? {
                            ...prev,
                            notificationType: value as NotificationType,
                          }
                        : null,
                    )
                  }
                  options={notificationOptions}
                  description={
                    strings.ManageAlertsNotificationDescription ||
                    strings.CreateAlertNotificationDescription
                  }
                />
              </SharePointSection>
              <SharePointSection
                title={strings.CreateAlertActionLinkSectionTitle}
              >
                <SharePointInput
                  label={strings.CreateAlertLinkUrlLabel}
                  value={editingAlert.linkUrl || ""}
                  onChange={(value) => {
                    setEditingAlert((prev) =>
                      prev ? { ...prev, linkUrl: value } : null,
                    );
                    if (editErrors.linkUrl)
                      setEditErrors((prev) => ({
                        ...prev,
                        linkUrl: undefined,
                      }));
                  }}
                  placeholder={strings.CreateAlertLinkUrlPlaceholder}
                  error={editErrors.linkUrl}
                  description={strings.CreateAlertLinkUrlDescription}
                />

                {editingAlert.linkUrl && !useMultiLanguage && (
                  <SharePointInput
                    label={strings.CreateAlertLinkDescriptionLabel}
                    value={editingAlert.linkDescription || ""}
                    onChange={(value) => {
                      setEditingAlert((prev) =>
                        prev ? { ...prev, linkDescription: value } : null,
                      );
                      if (editErrors.linkDescription)
                        setEditErrors((prev) => ({
                          ...prev,
                          linkDescription: undefined,
                        }));
                    }}
                    placeholder={strings.CreateAlertLinkDescriptionPlaceholder}
                    required={!!editingAlert.linkUrl}
                    error={editErrors.linkDescription}
                    description={strings.CreateAlertLinkDescriptionDescription}
                  />
                )}

                {editingAlert.linkUrl && useMultiLanguage && (
                  <div className={styles.infoMessage}>
                    <p>
                      💡 <strong>Multi-Language Mode:</strong> Link descriptions
                      are configured per language in the Multi-Language Content
                      section above.
                    </p>
                  </div>
                )}
              </SharePointSection>
              <SharePointSection
                title={strings.CreateAlertTargetSitesSectionTitle}
              >
                <SiteSelector
                  selectedSites={editingAlert.targetSites || []}
                  onSitesChange={(sites: string[]) => {
                    setEditingAlert((prev) =>
                      prev ? { ...prev, targetSites: sites } : null,
                    );
                    if (editErrors.targetSites)
                      setEditErrors((prev) => ({
                        ...prev,
                        targetSites: undefined,
                      }));
                  }}
                  siteDetector={siteDetector}
                  graphClient={graphClient}
                />
                {editErrors.targetSites && (
                  <div className={styles.errorMessage}>
                    {editErrors.targetSites}
                  </div>
                )}
                <div className={styles.fieldDescription}>
                  {strings.ManageAlertsTargetSitesDescription ||
                    "Choose which SharePoint sites will display this alert."}
                </div>
              </SharePointSection>
              <SharePointSection
                title={strings.CreateAlertSchedulingSectionTitle}
              >
                <div className={styles.schedulingContainer}>
                  <div className={styles.schedulingHeader}>
                    <p className={styles.schedulingDescription}>
                      {strings.ManageAlertsSchedulingDescription ||
                        "Configure when this alert will be visible to users. Leave dates empty for immediate activation and manual control."}
                    </p>
                  </div>

                  <div className={styles.schedulingGrid}>
                    <div>
                      <SharePointInput
                        label={strings.CreateAlertStartDateLabel}
                        type="datetime-local"
                        value={DateUtils.toDateTimeLocalValue(
                          editingAlert.scheduledStart,
                        )}
                        onChange={(value) => {
                          setEditingAlert((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  scheduledStart: value
                                    ? new Date(value)
                                    : undefined,
                                }
                              : null,
                          );
                          if (editErrors.scheduledStart)
                            setEditErrors((prev) => ({
                              ...prev,
                              scheduledStart: undefined,
                            }));
                        }}
                        error={editErrors.scheduledStart}
                        description={
                          strings.ManageAlertsStartDateDescription ||
                          strings.CreateAlertStartDateDescription
                        }
                      />
                    </div>

                    <div>
                      <SharePointInput
                        label={strings.CreateAlertEndDateLabel}
                        type="datetime-local"
                        value={DateUtils.toDateTimeLocalValue(
                          editingAlert.scheduledEnd,
                        )}
                        onChange={(value) => {
                          setEditingAlert((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  scheduledEnd: value
                                    ? new Date(value)
                                    : undefined,
                                }
                              : null,
                          );
                          if (editErrors.scheduledEnd)
                            setEditErrors((prev) => ({
                              ...prev,
                              scheduledEnd: undefined,
                            }));
                        }}
                        error={editErrors.scheduledEnd}
                        description={
                          strings.ManageAlertsEndDateDescription ||
                          strings.CreateAlertEndDateDescription
                        }
                      />
                    </div>
                  </div>

                  {/* Schedule Summary */}
                  <div className={styles.scheduleSummary}>
                    <h4>
                      {strings.ManageAlertsScheduleSummaryTitle ||
                        "Schedule Summary"}
                    </h4>
                    {!editingAlert.scheduledStart &&
                    !editingAlert.scheduledEnd ? (
                      <p>
                        {strings.ManageAlertsScheduleImmediate ||
                          "⚡ Alert is active immediately and requires manual deactivation."}
                      </p>
                    ) : editingAlert.scheduledStart &&
                      !editingAlert.scheduledEnd ? (
                      <p>
                        {Text.format(
                          strings.ManageAlertsScheduleStartOnly ||
                            "📅 Alert activates on {0} and remains active until manually deactivated.",
                          new Date(
                            editingAlert.scheduledStart,
                          ).toLocaleString(),
                        )}
                      </p>
                    ) : !editingAlert.scheduledStart &&
                      editingAlert.scheduledEnd ? (
                      <p>
                        {Text.format(
                          strings.ManageAlertsScheduleEndOnly ||
                            "⏰ Alert is active immediately until {0}.",
                          new Date(editingAlert.scheduledEnd).toLocaleString(),
                        )}
                      </p>
                    ) : (
                      <p>
                        {Text.format(
                          strings.ManageAlertsScheduleWindow ||
                            "📅 Active from {0} to {1}.",
                          new Date(
                            editingAlert.scheduledStart!,
                          ).toLocaleString(),
                          new Date(editingAlert.scheduledEnd!).toLocaleString(),
                        )}
                      </p>
                    )}
                  </div>

                  {/* Time Zone Info */}
                  <div className={styles.timezoneInfo}>
                    <p>
                      {Text.format(
                        strings.ManageAlertsScheduleTimezone ||
                          "🌍 All times are in your local time zone ({0}).",
                        Intl.DateTimeFormat().resolvedOptions().timeZone,
                      )}
                    </p>
                  </div>
                </div>
              </SharePointSection>
              {/* Action Buttons */}
              <div className={styles.formActions}>
                <SharePointButton
                  variant="primary"
                  onClick={handleSaveEdit}
                  disabled={isEditingAlert}
                  icon={<Save24Regular />}
                >
                  {isEditingAlert
                    ? strings.ManageAlertsSavingChangesLabel ||
                      "Saving Changes..."
                    : strings.ManageAlertsSaveChangesButton || "Save Changes"}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={handleCancelEdit}
                  disabled={isEditingAlert}
                >
                  {strings.Cancel}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={() => setShowPreview(!showPreview)}
                  icon={<Eye24Regular />}
                >
                  {showPreview
                    ? strings.CreateAlertHidePreview
                    : strings.CreateAlertShowPreview}
                </SharePointButton>
              </div>
            </div>

            {/* Preview Column */}
            {showPreview && (
              <div className={styles.formColumn}>
                <div className={styles.alertCard}>
                  <h3>
                    {strings.ManageAlertsLivePreviewTitle || "Live Preview"}
                  </h3>

                  {/* Multi-language preview mode selector */}
                  {useMultiLanguage &&
                    editingAlert.languageContent &&
                    editingAlert.languageContent.length > 0 && (
                      <div className={styles.previewLanguageSelector}>
                        <label className={styles.previewLabel}>
                          {strings.CreateAlertPreviewLanguageLabel}
                        </label>
                        <div className={styles.previewLanguageButtons}>
                          {editingAlert.languageContent.map(
                            (content, index) => {
                              const lang = supportedLanguages.find(
                                (l) => l.code === content.language,
                              );
                              return (
                                <SharePointButton
                                  key={content.language}
                                  variant={
                                    index === 0 ? "primary" : "secondary"
                                  }
                                  onClick={() => {
                                    // Move selected language to front for preview
                                    const reorderedContent = [
                                      content,
                                      ...editingAlert.languageContent!.filter(
                                        (_, i) => i !== index,
                                      ),
                                    ];
                                    setEditingAlert((prev) =>
                                      prev
                                        ? {
                                            ...prev,
                                            languageContent: reorderedContent,
                                          }
                                        : null,
                                    );
                                  }}
                                  className={styles.previewLanguageButton}
                                >
                                  {lang?.flag || content.language}{" "}
                                  {lang?.nativeName || content.language}
                                </SharePointButton>
                              );
                            },
                          )}
                        </div>
                      </div>
                    )}

                  <AlertPreview
                    title={
                      useMultiLanguage &&
                      editingAlert.languageContent &&
                      editingAlert.languageContent.length > 0
                        ? editingAlert.languageContent[0]?.title ||
                          strings.CreateAlertMultiLanguagePreviewTitle
                        : editingAlert.title || strings.AlertPreviewDefaultTitle
                    }
                    description={
                      useMultiLanguage &&
                      editingAlert.languageContent &&
                      editingAlert.languageContent.length > 0
                        ? editingAlert.languageContent[0]?.description ||
                          strings.CreateAlertMultiLanguagePreviewDescription
                        : editingAlert.description ||
                          strings.AlertPreviewDefaultDescription
                    }
                    alertType={
                      getCurrentAlertType() || {
                        name: strings.AlertTypeInfo,
                        iconName: "Info",
                        backgroundColor: "#0078d4",
                        textColor: "#ffffff",
                        additionalStyles: "",
                        priorityStyles: {},
                      }
                    }
                    priority={editingAlert.priority}
                    linkUrl={editingAlert.linkUrl}
                    linkDescription={
                      useMultiLanguage &&
                      editingAlert.languageContent &&
                      editingAlert.languageContent.length > 0
                        ? editingAlert.languageContent[0]?.linkDescription ||
                          strings.CreateAlertLinkPreviewFallback
                        : editingAlert.linkDescription ||
                          strings.CreateAlertLinkPreviewFallback
                    }
                    isPinned={editingAlert.isPinned}
                  />

                  {/* Multi-language preview info */}
                  {useMultiLanguage &&
                    editingAlert.languageContent &&
                    editingAlert.languageContent.length > 0 && (
                      <div className={styles.multiLanguagePreviewInfo}>
                        <p>
                          <strong>
                            {strings.ManageAlertsMultiLanguagePreviewHeading ||
                              "Multi-language Preview"}
                          </strong>
                        </p>
                        <p>
                          {Text.format(
                            strings.ManageAlertsMultiLanguagePreviewCurrentLanguage ||
                              "Currently previewing: {0}",
                            supportedLanguages.find(
                              (l) =>
                                l.code ===
                                editingAlert.languageContent![0]?.language,
                            )?.nativeName ||
                              editingAlert.languageContent![0]?.language,
                          )}
                        </p>
                        <p>
                          {Text.format(
                            strings.ManageAlertsMultiLanguagePreviewAvailableLanguages ||
                              "Available in {0} language(s): {1}",
                            editingAlert.languageContent.length,
                            editingAlert.languageContent
                              .map(
                                (c) =>
                                  supportedLanguages.find(
                                    (l) => l.code === c.language,
                                  )?.flag || c.language,
                              )
                              .join(" "),
                          )}
                        </p>
                      </div>
                    )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
      {dialogs}
    </>
  );
};

export default ManageAlertsTab;
