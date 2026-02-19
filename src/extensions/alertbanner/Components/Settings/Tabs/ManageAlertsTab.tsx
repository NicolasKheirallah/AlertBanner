import * as React from "react";
import {
  Delete24Regular,
  Edit24Regular,
  Save24Regular,
  Filter24Regular,
  ChevronDown24Regular,
  ChevronUp24Regular,
} from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  ISharePointSelectOption,
} from "../../UI/SharePointControls";
import AttachmentManager from "../../UI/AttachmentManager";
import {
  AlertPriority,
  NotificationType,
  IAlertType,
  TargetLanguage,
  ContentType,
  IPersonField,
  TranslationStatus,
  ILanguageContent,
} from "../../Alerts/IAlerts";
import { AlertSortMode } from "../../Context/AlertsContext";
import {
  LanguageAwarenessService,
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
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "../AlertSettings.module.scss";
import { validateAlertData } from "../../Utils/AlertValidation";
import { getLocalizedValidationMessage } from "../../Utils/AlertValidationLocalization";
import { useLanguageOptions } from "../../Hooks/useLanguageOptions";
import { usePriorityOptions } from "../../Hooks/usePriorityOptions";
import { useFluentDialogs } from "../../Hooks/useFluentDialogs";
import { CopilotService } from "../../Services/CopilotService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import AlertEditorForm from "./AlertEditorForm";
import EmptyState from "../../UI/EmptyState";
import AlertListShimmer from "../../UI/AlertListShimmer";
import { AlertMutationService } from "../../Services/AlertMutationService";
import ManageAlertCard, { IDisplayAlertItem } from "./ManageAlertCard";

interface IManageFiltersState {
  contentTypeFilter: "all" | ContentType;
  priorityFilter: "all" | AlertPriority;
  alertTypeFilter: "all" | string;
  statusFilter: "all" | string;
  contentStatusFilter: "all" | ContentStatus;
  languageFilter: "all" | TargetLanguage;
  notificationFilter: "all" | NotificationType;
  dateFilter: "all" | "today" | "week" | "month" | "custom";
  customDateFrom: string;
  customDateTo: string;
  searchTerm: string;
  showFilters: boolean;
}

const DEFAULT_MANAGE_FILTERS: IManageFiltersState = {
  contentTypeFilter: ContentType.Alert,
  priorityFilter: "all",
  alertTypeFilter: "all",
  statusFilter: "all",
  contentStatusFilter: "all",
  languageFilter: "all",
  notificationFilter: "all",
  dateFilter: "all",
  customDateFrom: "",
  customDateTo: "",
  searchTerm: "",
  showFilters: false,
};

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
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  copilotEnabled?: boolean;
  onDirtyStateChange?: (hasUnsavedChanges: boolean) => void;
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
  userTargetingEnabled,
  notificationsEnabled,
  enableTargetSite,
  copilotEnabled,
  onDirtyStateChange,
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
  const [isBulkActionInFlight, setIsBulkActionInFlight] = React.useState(false);
  const [busyAlertIds, setBusyAlertIds] = React.useState<string[]>([]);
  const advancedFiltersId = "manage-alerts-advanced-filters";
  
  // Drag and drop state
  const [draggingAlertId, setDraggingAlertId] = React.useState<string | null>(null);
  const [dragOverAlertId, setDragOverAlertId] = React.useState<string | null>(null);
  
  // Sort mode for Manage Alerts tab (default to priority)
  const [manageSortMode, setManageSortMode] = React.useState<"priority" | "manual">("priority");
  const [alertsListId, setAlertsListId] = React.useState<string>("");
  const [editingAlertSiteId, setEditingAlertSiteId] =
    React.useState<string>("");
  const [editingAlertListId, setEditingAlertListId] =
    React.useState<string>("");
  const [copilotAvailability, setCopilotAvailability] = React.useState<
    "unknown" | "available" | "unavailable"
  >("unknown");
  const searchInputContainerRef = React.useRef<HTMLDivElement>(null);
  const initialEditSnapshotRef = React.useRef<string>("");
  const didHydrateFiltersRef = React.useRef(false);

  const filtersStorageKey = React.useMemo(() => {
    const siteId = context.pageContext.site.id
      .toString()
      .toLowerCase()
      .replace(/[{}]/g, "");
    return `alert-banner:manage-filters:${siteId}`;
  }, [context.pageContext.site.id]);

  const withBusyAlert = React.useCallback(
    async (alertId: string, operation: () => Promise<void>) => {
      setBusyAlertIds((prev) =>
        prev.includes(alertId) ? prev : [...prev, alertId],
      );
      try {
        await operation();
      } finally {
        setBusyAlertIds((prev) => prev.filter((id) => id !== alertId));
      }
    },
    [],
  );

  const captureEditSnapshot = React.useCallback(
    (value: IEditingAlert | null): string => {
      if (!value) {
        return "";
      }

      const normalizePeople = (items?: IPersonField[]): string[] =>
        (items || [])
          .map((item) =>
            (
              item.loginName ||
              item.email ||
              item.displayName ||
              ""
            ).toLowerCase(),
          )
          .filter((item) => item.length > 0)
          .sort();

      const normalizeLanguages = (items?: ILanguageContent[]) =>
        (items || [])
          .map((item) => ({
            language: item.language,
            title: (item.title || "").trim(),
            description: (item.description || "").trim(),
            linkDescription: (item.linkDescription || "").trim(),
            translationStatus: item.translationStatus,
          }))
          .sort((a, b) => a.language.localeCompare(b.language));

      return JSON.stringify({
        title: (value.title || "").trim(),
        description: (value.description || "").trim(),
        AlertType: value.AlertType,
        priority: value.priority,
        isPinned: value.isPinned,
        notificationType: value.notificationType,
        linkUrl: (value.linkUrl || "").trim(),
        linkDescription: (value.linkDescription || "").trim(),
        targetSites: (value.targetSites || []).slice().sort(),
        targetLanguage: value.targetLanguage,
        contentType: value.contentType,
        scheduledStart: value.scheduledStart
          ? new Date(value.scheduledStart).toISOString()
          : "",
        scheduledEnd: value.scheduledEnd
          ? new Date(value.scheduledEnd).toISOString()
          : "",
        targetUsers: normalizePeople(value.targetUsers),
        targetGroups: normalizePeople(value.targetGroups),
        languageGroup: value.languageGroup || "",
        languageContent: normalizeLanguages(value.languageContent),
      });
    },
    [],
  );

  const hasUnsavedEditChanges = React.useMemo(() => {
    if (!editingAlert) {
      return false;
    }
    return captureEditSnapshot(editingAlert) !== initialEditSnapshotRef.current;
  }, [captureEditSnapshot, editingAlert]);

  const resetFilters = React.useCallback(() => {
    setContentTypeFilter(DEFAULT_MANAGE_FILTERS.contentTypeFilter);
    setPriorityFilter(DEFAULT_MANAGE_FILTERS.priorityFilter);
    setAlertTypeFilter(DEFAULT_MANAGE_FILTERS.alertTypeFilter);
    setStatusFilter(DEFAULT_MANAGE_FILTERS.statusFilter);
    setLanguageFilter(DEFAULT_MANAGE_FILTERS.languageFilter);
    setNotificationFilter(DEFAULT_MANAGE_FILTERS.notificationFilter);
    setDateFilter(DEFAULT_MANAGE_FILTERS.dateFilter);
    setCustomDateFrom(DEFAULT_MANAGE_FILTERS.customDateFrom);
    setCustomDateTo(DEFAULT_MANAGE_FILTERS.customDateTo);
    setSearchTerm(DEFAULT_MANAGE_FILTERS.searchTerm);
    setContentStatusFilter(DEFAULT_MANAGE_FILTERS.contentStatusFilter);
  }, []);

  const focusSearchInput = React.useCallback(() => {
    const input = searchInputContainerRef.current?.querySelector("input");
    if (input instanceof HTMLInputElement) {
      input.focus();
      input.select();
    }
  }, []);

  const isEditableTarget = React.useCallback((target: EventTarget | null) => {
    if (!(target instanceof HTMLElement)) {
      return false;
    }

    const tag = target.tagName.toLowerCase();
    return (
      target.isContentEditable ||
      tag === "input" ||
      tag === "textarea" ||
      tag === "select" ||
      target.closest('[contenteditable="true"]') !== null
    );
  }, []);

  const stripFilterEmoji = React.useCallback((value: string): string => {
    return value
      .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}\u200D]/gu, "")
      .replace(/\s{2,}/g, " ")
      .trim();
  }, []);

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

  React.useEffect(() => {
    try {
      const raw = localStorage.getItem(filtersStorageKey);
      if (!raw) {
        didHydrateFiltersRef.current = true;
        return;
      }

      const saved = JSON.parse(raw) as Partial<IManageFiltersState>;
      if (saved.contentTypeFilter !== undefined) {
        setContentTypeFilter(saved.contentTypeFilter);
      }
      if (saved.priorityFilter !== undefined) {
        setPriorityFilter(saved.priorityFilter);
      }
      if (saved.alertTypeFilter !== undefined) {
        setAlertTypeFilter(saved.alertTypeFilter);
      }
      if (saved.statusFilter !== undefined) {
        setStatusFilter(saved.statusFilter);
      }
      if (saved.contentStatusFilter !== undefined) {
        setContentStatusFilter(saved.contentStatusFilter);
      }
      if (saved.languageFilter !== undefined) {
        setLanguageFilter(saved.languageFilter);
      }
      if (saved.notificationFilter !== undefined) {
        setNotificationFilter(saved.notificationFilter);
      }
      if (saved.dateFilter !== undefined) {
        setDateFilter(saved.dateFilter);
      }
      if (saved.customDateFrom !== undefined) {
        setCustomDateFrom(saved.customDateFrom);
      }
      if (saved.customDateTo !== undefined) {
        setCustomDateTo(saved.customDateTo);
      }
      if (saved.searchTerm !== undefined) {
        setSearchTerm(saved.searchTerm);
      }
      if (saved.showFilters !== undefined) {
        setShowFilters(saved.showFilters);
      }
    } catch (error) {
      logger.warn("ManageAlertsTab", "Failed to hydrate saved filters", error);
    } finally {
      didHydrateFiltersRef.current = true;
    }
  }, [filtersStorageKey]);

  React.useEffect(() => {
    if (!didHydrateFiltersRef.current) {
      return;
    }

    const payload: IManageFiltersState = {
      contentTypeFilter,
      priorityFilter,
      alertTypeFilter,
      statusFilter,
      contentStatusFilter,
      languageFilter,
      notificationFilter,
      dateFilter,
      customDateFrom,
      customDateTo,
      searchTerm,
      showFilters,
    };

    try {
      localStorage.setItem(filtersStorageKey, JSON.stringify(payload));
    } catch (error) {
      logger.warn("ManageAlertsTab", "Failed to persist filters", error);
    }
  }, [
    alertTypeFilter,
    contentStatusFilter,
    contentTypeFilter,
    customDateFrom,
    customDateTo,
    dateFilter,
    filtersStorageKey,
    languageFilter,
    notificationFilter,
    priorityFilter,
    searchTerm,
    showFilters,
    statusFilter,
  ]);

  React.useEffect(() => {
    if (!onDirtyStateChange) {
      return;
    }

    onDirtyStateChange(!!editingAlert && hasUnsavedEditChanges);
  }, [editingAlert, hasUnsavedEditChanges, onDirtyStateChange]);

  React.useEffect(() => {
    const handleBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!editingAlert || !hasUnsavedEditChanges) {
        return;
      }

      event.preventDefault();
      event.returnValue = "";
    };

    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [editingAlert, hasUnsavedEditChanges]);

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

  React.useEffect(() => {
    let isMounted = true;

    if (!copilotEnabled) {
      setCopilotAvailability("unknown");
      return () => {
        isMounted = false;
      };
    }

    setCopilotAvailability("unknown");

    copilotService
      .checkAccess()
      .then((isAvailable) => {
        if (!isMounted) {
          return;
        }

        setCopilotAvailability(isAvailable ? "available" : "unavailable");
      })
      .catch((error) => {
        logger.warn(
          "ManageAlertsTab",
          "Copilot access preflight failed",
          error,
        );
        if (!isMounted) {
          return;
        }

        setCopilotAvailability("unavailable");
      });

    return () => {
      isMounted = false;
      copilotService.cancelActiveOperation();
    };
  }, [copilotEnabled, copilotService]);

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
        label: Text.format(
          strings.ManageAlertsAlertTypeFromSourceSiteLabel,
          editingAlert.AlertType,
        ),
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
    if (selectedAlerts.length === 0 || isBulkActionInFlight) return;

    const shouldDelete = await confirm({
      title: strings.ManageAlertsBulkDeleteDialogTitle,
      message: Text.format(
        strings.ManageAlertsBulkDeleteDialogMessage,
        selectedAlerts.length,
      ),
      confirmText: strings.Delete,
    });
    if (!shouldDelete) {
      return;
    }

    const selectedSet = new Set(selectedAlerts);
    const previousAlerts = existingAlerts;
    setIsBulkActionInFlight(true);

    // Optimistic remove for faster perceived response.
    setExistingAlerts((prev) =>
      prev.filter((item) => !selectedSet.has(item.id)),
    );
    setSelectedAlerts([]);

    try {
      await Promise.all(
        selectedAlerts.map((alertId) => alertService.deleteAlert(alertId)),
      );

      await loadExistingAlerts();
      notificationService.showSuccess(
        Text.format(
          strings.ManageAlertsBulkDeleteSuccessMessage,
          selectedAlerts.length,
        ),
        strings.ManageAlertsBulkDeleteSuccessTitle,
      );
    } catch (error) {
      setExistingAlerts(previousAlerts);
      setSelectedAlerts(selectedAlerts);
      logger.error("ManageAlertsTab", "Error deleting alerts", error);
      notificationService.showError(
        strings.ManageAlertsBulkDeleteErrorMessage,
        strings.ManageAlertsBulkDeleteErrorTitle,
      );
    } finally {
      setIsBulkActionInFlight(false);
    }
  }, [
    alertService,
    confirm,
    existingAlerts,
    isBulkActionInFlight,
    loadExistingAlerts,
    notificationService,
    selectedAlerts,
    setExistingAlerts,
    setSelectedAlerts,
  ]);

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
          initialEditSnapshotRef.current = captureEditSnapshot(editingData);
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
          initialEditSnapshotRef.current = captureEditSnapshot(editingData);
          setUseMultiLanguage(false);
        }

        setShowPreview(true); // Ensure preview is visible when opening edit
      } catch (error) {
        logger.error("ManageAlertsTab", "Error opening edit dialog", error);
        notificationService.showError(
          strings.ManageAlertsEditOpenErrorMessage,
          strings.ManageAlertsEditOpenErrorTitle,
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
      captureEditSnapshot,
    ],
  );

  const handleDeleteAlert = React.useCallback(
    async (alertId: string, alertTitle: string) => {
      if (isBulkActionInFlight || busyAlertIds.includes(alertId)) {
        return;
      }

      const shouldDelete = await confirm({
        title: strings.ManageAlertsDeleteDialogTitle,
        message: Text.format(
          strings.ManageAlertsDeleteDialogMessage,
          alertTitle,
        ),
        confirmText: strings.Delete,
      });
      if (!shouldDelete) {
        return;
      }

      await withBusyAlert(alertId, async () => {
        const previousAlerts = existingAlerts;
        setExistingAlerts((prev) => prev.filter((item) => item.id !== alertId));
        setSelectedAlerts((prev) => prev.filter((id) => id !== alertId));

        try {
          const alert = existingAlerts.find((a) => a.id === alertId);
          const folderName = alert?.languageGroup || alertTitle;
          const alertSiteId = alert
            ? alertService.getAlertSiteId(alert.id)
            : undefined;

          await alertService.deleteAlert(alertId);
          await imageStorageService.deleteImageFolder(folderName, alertSiteId);
          await loadExistingAlerts();

          notificationService.showSuccess(
            Text.format(strings.ManageAlertsDeleteSuccessMessage, alertTitle),
            strings.ManageAlertsDeleteSuccessTitle,
          );
        } catch (error) {
          setExistingAlerts(previousAlerts);
          logger.error("ManageAlertsTab", "Error deleting alert", error);
          notificationService.showError(
            strings.ManageAlertsDeleteErrorMessage,
            strings.ManageAlertsDeleteErrorTitle,
          );
        }
      });
    },
    [
      alertService,
      busyAlertIds,
      existingAlerts,
      imageStorageService,
      isBulkActionInFlight,
      loadExistingAlerts,
      notificationService,
      confirm,
      setExistingAlerts,
      setSelectedAlerts,
      withBusyAlert,
    ],
  );

  const handlePublishDraft = React.useCallback(
    async (draft: IAlertItem) => {
      if (isBulkActionInFlight || busyAlertIds.includes(draft.id)) {
        return;
      }

      const shouldPublish = await confirm({
        title: strings.ManageAlertsPublishDialogTitle,
        message: Text.format(
          strings.ManageAlertsPublishDialogMessage,
          draft.title,
        ),
        confirmText: strings.ManageAlertsPublishDialogConfirm,
      });
      if (!shouldPublish) {
        return;
      }

      await withBusyAlert(draft.id, async () => {
        const previousAlerts = existingAlerts;
        setExistingAlerts((prev) =>
          prev.map((item) =>
            item.id === draft.id
              ? {
                  ...item,
                  contentType: ContentType.Alert,
                  status: "Active",
                }
              : item,
          ),
        );

        try {
          const publishedAlert: Partial<IAlertItem> = {
            ...draft,
            contentType: ContentType.Alert,
            status: "Active",
          };

          await alertService.updateAlert(draft.id, publishedAlert);
          await loadExistingAlerts();
          notificationService.showSuccess(
            Text.format(strings.ManageAlertsPublishSuccessMessage, draft.title),
            strings.ManageAlertsPublishSuccessTitle,
          );
        } catch (error) {
          setExistingAlerts(previousAlerts);
          logger.error("ManageAlertsTab", "Error publishing draft", error);
          notificationService.showError(
            strings.ManageAlertsPublishErrorMessage,
            strings.ManageAlertsPublishErrorTitle,
          );
        }
      });
    },
    [
      alertService,
      busyAlertIds,
      confirm,
      existingAlerts,
      isBulkActionInFlight,
      loadExistingAlerts,
      notificationService,
      setExistingAlerts,
      withBusyAlert,
    ],
  );

  const handleSubmitForReview = React.useCallback(
    async (alert: IAlertItem) => {
      if (isBulkActionInFlight || busyAlertIds.includes(alert.id)) {
        return;
      }

      const shouldSubmit = await confirm({
        title: strings.ManageAlertsSubmitDialogTitle,
        message: Text.format(
          strings.ManageAlertsSubmitDialogMessage,
          alert.title,
        ),
        confirmText: strings.ManageAlertsSubmitDialogConfirm,
      });
      if (!shouldSubmit) {
        return;
      }

      await withBusyAlert(alert.id, async () => {
        const previousAlerts = existingAlerts;
        setExistingAlerts((prev) =>
          prev.map((item) =>
            item.id === alert.id
              ? { ...item, contentStatus: ContentStatus.PendingReview }
              : item,
          ),
        );

        try {
          await alertService.submitAlert(alert.id);
          await loadExistingAlerts();
          notificationService.showSuccess(
            Text.format(strings.ManageAlertsSubmitSuccessMessage, alert.title),
            strings.ManageAlertsSubmitSuccessTitle,
          );
        } catch (error) {
          setExistingAlerts(previousAlerts);
          logger.error("ManageAlertsTab", "Error submitting alert", error);
          notificationService.showError(
            strings.ManageAlertsSubmitErrorMessage,
            strings.ManageAlertsSubmitErrorTitle,
          );
        }
      });
    },
    [
      alertService,
      busyAlertIds,
      confirm,
      existingAlerts,
      isBulkActionInFlight,
      loadExistingAlerts,
      notificationService,
      setExistingAlerts,
      withBusyAlert,
    ],
  );

  const handleApprove = React.useCallback(
    async (alert: IAlertItem) => {
      if (isBulkActionInFlight || busyAlertIds.includes(alert.id)) {
        return;
      }

      const shouldApprove = await confirm({
        title: strings.ManageAlertsApproveDialogTitle,
        message: Text.format(
          strings.ManageAlertsApproveDialogMessage,
          alert.title,
        ),
        confirmText: strings.ManageAlertsApproveDialogConfirm,
      });
      if (!shouldApprove) {
        return;
      }

      await withBusyAlert(alert.id, async () => {
        const previousAlerts = existingAlerts;
        setExistingAlerts((prev) =>
          prev.map((item) =>
            item.id === alert.id
              ? { ...item, contentStatus: ContentStatus.Approved }
              : item,
          ),
        );

        try {
          await alertService.approveAlert(alert.id);
          await loadExistingAlerts();
          notificationService.showSuccess(
            Text.format(strings.ManageAlertsApproveSuccessMessage, alert.title),
            strings.ManageAlertsApproveSuccessTitle,
          );
        } catch (error) {
          setExistingAlerts(previousAlerts);
          logger.error("ManageAlertsTab", "Error approving alert", error);
          notificationService.showError(
            strings.ManageAlertsApproveErrorMessage,
            strings.ManageAlertsApproveErrorTitle,
          );
        }
      });
    },
    [
      alertService,
      busyAlertIds,
      confirm,
      existingAlerts,
      isBulkActionInFlight,
      loadExistingAlerts,
      notificationService,
      setExistingAlerts,
      withBusyAlert,
    ],
  );

  const handleReject = React.useCallback(
    async (alert: IAlertItem) => {
      if (isBulkActionInFlight || busyAlertIds.includes(alert.id)) {
        return;
      }

      const notes = await prompt({
        title: strings.ManageAlertsRejectDialogTitle,
        message: Text.format(
          strings.ManageAlertsRejectDialogMessage,
          alert.title,
        ),
        confirmText: strings.ManageAlertsRejectDialogConfirm,
        label: strings.ManageAlertsRejectDialogLabel,
        placeholder: strings.ManageAlertsRejectDialogPlaceholder,
      });
      if (notes === null) return;

      await withBusyAlert(alert.id, async () => {
        const previousAlerts = existingAlerts;
        setExistingAlerts((prev) =>
          prev.map((item) =>
            item.id === alert.id
              ? { ...item, contentStatus: ContentStatus.Rejected }
              : item,
          ),
        );

        try {
          await alertService.rejectAlert(alert.id, notes);
          await loadExistingAlerts();
          notificationService.showSuccess(
            Text.format(strings.ManageAlertsRejectSuccessMessage, alert.title),
            strings.ManageAlertsRejectSuccessTitle,
          );
        } catch (error) {
          setExistingAlerts(previousAlerts);
          logger.error("ManageAlertsTab", "Error rejecting alert", error);
          notificationService.showError(
            strings.ManageAlertsRejectErrorMessage,
            strings.ManageAlertsRejectErrorTitle,
          );
        }
      });
    },
    [
      alertService,
      busyAlertIds,
      existingAlerts,
      isBulkActionInFlight,
      loadExistingAlerts,
      notificationService,
      prompt,
      setExistingAlerts,
      withBusyAlert,
    ],
  );

  // Drag and Drop Handlers
  const handleDragStart = React.useCallback((alertId: string) => {
    setDraggingAlertId(alertId);
  }, []);

  const handleDragEnd = React.useCallback(() => {
    setDraggingAlertId(null);
    setDragOverAlertId(null);
  }, []);

  const handleDragOver = React.useCallback((alertId: string) => {
    if (draggingAlertId && draggingAlertId !== alertId) {
      setDragOverAlertId(alertId);
    }
  }, [draggingAlertId]);

  const handleDrop = React.useCallback(async (targetAlertId: string) => {
    if (!draggingAlertId || draggingAlertId === targetAlertId) {
      return;
    }

    // Find indices in existingAlerts (the source of truth)
    const draggedIndex = existingAlerts.findIndex(a => a.id === draggingAlertId);
    const targetIndex = existingAlerts.findIndex(a => a.id === targetAlertId);
    
    if (draggedIndex === -1 || targetIndex === -1) return;

    // Reorder the alerts
    const newAlerts = [...existingAlerts];
    const [removed] = newAlerts.splice(draggedIndex, 1);
    newAlerts.splice(targetIndex, 0, removed);

    // Update sort orders
    const updatedAlerts = newAlerts.map((alert, index) => ({
      ...alert,
      sortOrder: index,
    }));

    // Update local state immediately for smooth UX
    setExistingAlerts(updatedAlerts);

    // Persist to SharePoint (debounced - only update the changed alerts)
    try {
      // Only update the dragged alert and the alerts between dragged and target positions
      const minIndex = Math.min(draggedIndex, targetIndex);
      const maxIndex = Math.max(draggedIndex, targetIndex);
      const alertOrders = updatedAlerts
        .slice(minIndex, maxIndex + 1)
        .map((alert) => ({
          alertId: alert.id,
          sortOrder: alert.sortOrder ?? 0,
        }));
      await alertService.updateAlertsSortOrder(alertOrders);
      notificationService.showSuccess("Alert order updated", "Success");
    } catch (error) {
      logger.error("ManageAlertsTab", "Error updating sort order", error);
      notificationService.showError("Failed to save alert order", "Error");
      // Reload to get correct state
      loadExistingAlerts();
    }

    setDraggingAlertId(null);
    setDragOverAlertId(null);
  }, [draggingAlertId, existingAlerts, alertService, notificationService, setExistingAlerts, loadExistingAlerts]);

  // Sort Order Input Handler
  const handleSortOrderChange = React.useCallback(async (alertId: string, newOrder: number) => {
    const alertIndex = existingAlerts.findIndex(a => a.id === alertId);
    if (alertIndex === -1) return;

    // Clamp newOrder to valid range
    const clampedOrder = Math.max(0, Math.min(newOrder, existingAlerts.length - 1));
    if (clampedOrder === alertIndex) return; // No change needed

    // Create new array with reordered alerts
    const newAlerts = [...existingAlerts];
    const [movedAlert] = newAlerts.splice(alertIndex, 1);
    newAlerts.splice(clampedOrder, 0, movedAlert);

    // Reindex all alerts
    const reindexedAlerts = newAlerts.map((alert, index) => ({
      ...alert,
      sortOrder: index,
    }));

    setExistingAlerts(reindexedAlerts);

    // Persist to SharePoint (only update affected range)
    try {
      const minIndex = Math.min(alertIndex, clampedOrder);
      const maxIndex = Math.max(alertIndex, clampedOrder);
      const alertOrders = reindexedAlerts
        .slice(minIndex, maxIndex + 1)
        .map((alert) => ({
          alertId: alert.id,
          sortOrder: alert.sortOrder ?? 0,
        }));
      await alertService.updateAlertsSortOrder(alertOrders);
      notificationService.showSuccess("Alert order updated", "Success");
    } catch (error) {
      logger.error("ManageAlertsTab", "Error updating sort order", error);
      notificationService.showError("Failed to save sort order", "Error");
      loadExistingAlerts();
    }
  }, [existingAlerts, alertService, notificationService, setExistingAlerts, loadExistingAlerts]);

  // Move to Top/Bottom Handlers
  const handleMoveToTop = React.useCallback(async (alertId: string) => {
    await handleSortOrderChange(alertId, 0);
  }, [handleSortOrderChange]);

  const handleMoveToBottom = React.useCallback(async (alertId: string) => {
    await handleSortOrderChange(alertId, Number.MAX_SAFE_INTEGER); // Will be clamped to valid range
  }, [handleSortOrderChange]);

  // Validation using shared utility with localization
  const validateEditForm = React.useCallback((): boolean => {
    if (!editingAlert) return false;

    const validationData: Partial<IAlertItem> = {
      ...editingAlert,
      scheduledStart: editingAlert.scheduledStart
        ? editingAlert.scheduledStart.toISOString()
        : undefined,
      scheduledEnd: editingAlert.scheduledEnd
        ? editingAlert.scheduledEnd.toISOString()
        : undefined,
    };

    const validationErrors = validateAlertData(validationData, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage,
      getString: getLocalizedValidationMessage,
      validateTargetSites: enableTargetSite,
    });

    setEditErrors(validationErrors);
    return Object.keys(validationErrors).length === 0;
  }, [
    editingAlert,
    useMultiLanguage,
    languagePolicy,
    tenantDefaultLanguage,
    enableTargetSite,
  ]);

  const handleSaveEdit = React.useCallback(async () => {
    if (!editingAlert || !validateEditForm()) return;

    setIsEditingAlert(true);
    try {
      if (
        useMultiLanguage &&
        editingAlert.languageContent &&
        editingAlert.languageContent.length > 0
      ) {
        await AlertMutationService.syncMultiLanguageAlerts(
          alertService,
          editingAlert,
          existingAlerts,
          languagePolicy,
        );
      } else {
        await AlertMutationService.updateSingleAlert(
          alertService,
          editingAlert,
        );
      }

      setEditingAlert(null);
      setEditErrors({});
      setUseMultiLanguage(false);
      initialEditSnapshotRef.current = "";
      await loadExistingAlerts();
      notificationService.showSuccess(
        strings.ManageAlertsUpdateSuccessMessage,
        strings.ManageAlertsUpdateSuccessTitle,
      );
    } catch (error) {
      logger.error("ManageAlertsTab", "Error updating alert", error);
      notificationService.showError(
        strings.ManageAlertsUpdateErrorMessage,
        strings.ManageAlertsUpdateErrorTitle,
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

  const handleCancelEdit = React.useCallback(async () => {
    if (hasUnsavedEditChanges) {
      const canDiscard = await confirm({
        title: strings.ManageAlertsUnsavedChangesTitle,
        message: strings.ManageAlertsUnsavedChangesMessage,
        confirmText: strings.ManageAlertsDiscardChangesButton,
        cancelText: strings.ManageAlertsKeepEditingButton,
      });

      if (!canDiscard) {
        return;
      }
    }

    setEditingAlert(null);
    setEditErrors({});
    setUseMultiLanguage(false);
    setSupplementalAlertTypes([]);
    initialEditSnapshotRef.current = "";
  }, [
    confirm,
    hasUnsavedEditChanges,
    setEditingAlert,
    setSupplementalAlertTypes,
  ]);

  const setEditingAlertData = React.useCallback(
    (value: React.SetStateAction<IEditingAlert>) => {
      setEditingAlert((prev) => {
        if (!prev) {
          return prev;
        }
        return typeof value === "function"
          ? (value as (prevState: IEditingAlert) => IEditingAlert)(prev)
          : value;
      });
    },
    [setEditingAlert],
  );

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

  // Group alerts by language group with enhanced filtering and sorting
  const groupedAlerts = React.useMemo(() => {
    let filteredAlerts = [...existingAlerts];

    // Apply sorting based on sort mode
    if (manageSortMode === "manual") {
      // Sort by sortOrder (nulls/undefined at the end), then by created date as tiebreaker
      filteredAlerts.sort((a, b) => {
        const orderA = typeof a.sortOrder === "number" ? a.sortOrder : Number.MAX_SAFE_INTEGER;
        const orderB = typeof b.sortOrder === "number" ? b.sortOrder : Number.MAX_SAFE_INTEGER;
        if (orderA !== orderB) {
          return orderA - orderB;
        }
        // Tiebreaker: created date (newest first)
        return new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime();
      });
    }
    // For "priority" mode, keep default order (which is by date from API)

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
    const displayItems: IDisplayAlertItem[] = [];

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
    manageSortMode,
  ]);

  const languageGroupCount = React.useMemo(
    () => groupedAlerts.filter((a) => a.isLanguageGroup).length,
    [groupedAlerts],
  );

  const contentTypeCounts = React.useMemo(() => {
    return existingAlerts.reduce(
      (acc, alert) => {
        acc.all += 1;
        if (alert.contentType === ContentType.Alert) {
          acc.alert += 1;
        } else if (alert.contentType === ContentType.Draft) {
          acc.draft += 1;
        } else if (alert.contentType === ContentType.Template) {
          acc.template += 1;
        }
        return acc;
      },
      { all: 0, alert: 0, draft: 0, template: 0 },
    );
  }, [existingAlerts]);

  const supportedLanguageMap = React.useMemo(() => {
    const map = new Map<string, ISupportedLanguage>();
    supportedLanguages.forEach((lang) => map.set(lang.code, lang));
    return map;
  }, [supportedLanguages]);

  const priorityLabelMap = React.useMemo(() => {
    const map = new Map<string, string>();
    priorityOptions.forEach((option) =>
      map.set(String(option.value), option.label),
    );
    return map;
  }, [priorityOptions]);

  const alertTypeStyleMap = React.useMemo(() => {
    const map = new Map<
      string,
      { backgroundColor?: string; textColor?: string }
    >();
    alertTypes.forEach((type) =>
      map.set(type.name, {
        backgroundColor: type.backgroundColor,
        textColor: type.textColor,
      }),
    );
    return map;
  }, [alertTypes]);

  const selectedAlertIds = React.useMemo(
    () => new Set(selectedAlerts),
    [selectedAlerts],
  );
  const hasActiveFilters = React.useMemo(
    () =>
      contentTypeFilter !== DEFAULT_MANAGE_FILTERS.contentTypeFilter ||
      priorityFilter !== DEFAULT_MANAGE_FILTERS.priorityFilter ||
      alertTypeFilter !== DEFAULT_MANAGE_FILTERS.alertTypeFilter ||
      statusFilter !== DEFAULT_MANAGE_FILTERS.statusFilter ||
      contentStatusFilter !== DEFAULT_MANAGE_FILTERS.contentStatusFilter ||
      languageFilter !== DEFAULT_MANAGE_FILTERS.languageFilter ||
      notificationFilter !== DEFAULT_MANAGE_FILTERS.notificationFilter ||
      dateFilter !== DEFAULT_MANAGE_FILTERS.dateFilter ||
      customDateFrom !== DEFAULT_MANAGE_FILTERS.customDateFrom ||
      customDateTo !== DEFAULT_MANAGE_FILTERS.customDateTo ||
      searchTerm.trim().length > 0,
    [
      alertTypeFilter,
      contentStatusFilter,
      contentTypeFilter,
      customDateFrom,
      customDateTo,
      dateFilter,
      languageFilter,
      notificationFilter,
      priorityFilter,
      searchTerm,
      statusFilter,
    ],
  );
  const activeFilterCount = React.useMemo(() => {
    let count = 0;
    if (contentTypeFilter !== DEFAULT_MANAGE_FILTERS.contentTypeFilter)
      count += 1;
    if (priorityFilter !== DEFAULT_MANAGE_FILTERS.priorityFilter) count += 1;
    if (alertTypeFilter !== DEFAULT_MANAGE_FILTERS.alertTypeFilter) count += 1;
    if (statusFilter !== DEFAULT_MANAGE_FILTERS.statusFilter) count += 1;
    if (contentStatusFilter !== DEFAULT_MANAGE_FILTERS.contentStatusFilter)
      count += 1;
    if (languageFilter !== DEFAULT_MANAGE_FILTERS.languageFilter) count += 1;
    if (notificationFilter !== DEFAULT_MANAGE_FILTERS.notificationFilter)
      count += 1;
    if (dateFilter !== DEFAULT_MANAGE_FILTERS.dateFilter) count += 1;
    if (customDateFrom !== DEFAULT_MANAGE_FILTERS.customDateFrom) count += 1;
    if (customDateTo !== DEFAULT_MANAGE_FILTERS.customDateTo) count += 1;
    if (searchTerm.trim().length > 0) count += 1;
    return count;
  }, [
    alertTypeFilter,
    contentStatusFilter,
    contentTypeFilter,
    customDateFrom,
    customDateTo,
    dateFilter,
    languageFilter,
    notificationFilter,
    priorityFilter,
    searchTerm,
    statusFilter,
  ]);

  // Load alerts on mount
  React.useEffect(() => {
    loadExistingAlerts();
  }, [loadExistingAlerts]);

  React.useEffect(() => {
    const existingIds = new Set(existingAlerts.map((alert) => alert.id));
    setSelectedAlerts((prev) => prev.filter((id) => existingIds.has(id)));
  }, [existingAlerts, setSelectedAlerts]);

  React.useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();
      const isMeta = event.ctrlKey || event.metaKey;
      const inEditable = isEditableTarget(event.target);

      if (isMeta && key === "f" && !editingAlert) {
        event.preventDefault();
        focusSearchInput();
        return;
      }

      if (
        isMeta &&
        key === "e" &&
        !editingAlert &&
        selectedAlerts.length === 1
      ) {
        event.preventDefault();
        const selectedAlert = groupedAlerts.find((alert) =>
          selectedAlerts.includes(alert.id),
        );
        if (selectedAlert) {
          handleEditAlert(selectedAlert);
        }
        return;
      }

      if (isMeta && key === "s" && editingAlert) {
        event.preventDefault();
        void handleSaveEdit();
        return;
      }

      if (event.key === "Escape") {
        if (editingAlert) {
          event.preventDefault();
          void handleCancelEdit();
          return;
        }

        if (selectedAlerts.length > 0 && !inEditable) {
          event.preventDefault();
          setSelectedAlerts([]);
        }
      }

      if (
        (event.key === "Delete" || event.key === "Backspace") &&
        !editingAlert &&
        !inEditable &&
        selectedAlerts.length > 0
      ) {
        event.preventDefault();
        void handleBulkDelete();
      }
    };

    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [
    editingAlert,
    focusSearchInput,
    groupedAlerts,
    handleBulkDelete,
    handleCancelEdit,
    handleEditAlert,
    handleSaveEdit,
    isEditableTarget,
    selectedAlerts,
    setSelectedAlerts,
  ]);

  return (
    <>
      {!editingAlert && (
        <div className={styles.tabPane}>
          <div className={styles.tabHeader}>
            <div>
              <h3>{strings.ManageAlerts}</h3>
              <p>{strings.ManageAlertsSubtitle}</p>
            </div>
            <div className={styles.flexRowGap12}>
              <SharePointButton
                variant="secondary"
                onClick={loadExistingAlerts}
                disabled={isLoadingAlerts || isBulkActionInFlight}
              >
                {isLoadingAlerts
                  ? strings.ManageAlertsRefreshingLabel
                  : strings.Refresh}
              </SharePointButton>
              <button
                type="button"
                className={styles.filterToggleButton}
                onClick={() => setShowFilters(!showFilters)}
                aria-expanded={showFilters}
                aria-controls={advancedFiltersId}
              >
                <Filter24Regular />
                {showFilters
                  ? strings.ManageAlertsHideFiltersLabel
                  : strings.ManageAlertsShowFiltersLabel}{" "}
                {strings.ManageAlertsFiltersLabel}
                {activeFilterCount > 0 && ` (${activeFilterCount})`}
                {showFilters ? (
                  <ChevronUp24Regular />
                ) : (
                  <ChevronDown24Regular />
                )}
              </button>
            </div>
          </div>

          {/* Sort Mode Toggle */}
          {!isLoadingAlerts && existingAlerts.length > 0 && (
            <div className={(styles as any).sortModeBar}>
              <span className={(styles as any).sortModeLabel}>Sort by:</span>
              <div className={(styles as any).sortModeToggle}>
                <button
                  type="button"
                  className={`${(styles as any).sortModeButton} ${manageSortMode === "priority" ? (styles as any).active : ""}`}
                  onClick={() => setManageSortMode("priority")}
                  title="Sort by priority"
                >
                  Priority
                </button>
                <button
                  type="button"
                  className={`${(styles as any).sortModeButton} ${manageSortMode === "manual" ? (styles as any).active : ""}`}
                  onClick={() => setManageSortMode("manual")}
                  title="Manual drag & drop sorting"
                >
                  Manual
                </button>
              </div>
              {manageSortMode === "manual" && (
                <span className={(styles as any).sortModeHint}>
                  Drag cards or use arrows to reorder
                </span>
              )}
            </div>
          )}

          {isLoadingAlerts ? (
            <AlertListShimmer />
          ) : existingAlerts.length === 0 ? (
            <EmptyState
              icon="AlertSolid"
              title={strings.ManageAlertsEmptyTitle}
              description={`${strings.ManageAlertsEmptyDescription}\n\n• ${strings.ManageAlertsEmptyReasonListsMissing}\n• ${strings.ManageAlertsEmptyReasonNoAccess}\n• ${strings.ManageAlertsEmptyReasonNoAlerts}`}
              actionText={strings.ManageAlertsEmptyCreateFirstButton}
              onAction={() => setActiveTab("create")}
            />
          ) : (
            <div className={styles.alertsList}>
              {selectedAlerts.length > 0 && (
                <div
                  className={styles.selectionActionBar}
                  role="region"
                  aria-live="polite"
                >
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
                      disabled={isLoadingAlerts || isBulkActionInFlight}
                    >
                      {strings.ManageAlertsEditSelectedButtonLabel}
                    </SharePointButton>
                  )}
                  <SharePointButton
                    variant="danger"
                    icon={<Delete24Regular />}
                    onClick={handleBulkDelete}
                    disabled={isLoadingAlerts || isBulkActionInFlight}
                  >
                    {Text.format(
                      strings.ManageAlertsDeleteSelectedButtonLabel,
                      selectedAlerts.length,
                    )}
                  </SharePointButton>
                  <SharePointButton
                    variant="secondary"
                    onClick={() => setSelectedAlerts([])}
                    disabled={isLoadingAlerts || isBulkActionInFlight}
                  >
                    {strings.Cancel}
                  </SharePointButton>
                </div>
              )}

              {/* Enhanced Filter Section */}
              <div className={styles.filterSection}>
                <div className={styles.filterHeader}>
                  <div className={styles.filterSummary}>
                    <span>
                      {Text.format(
                        strings.ManageAlertsFilterSummaryCount,
                        groupedAlerts.length,
                        existingAlerts.length,
                      )}
                    </span>
                    <span>
                      {Text.format(
                        strings.ManageAlertsFilterSummaryLanguageGroups,
                        languageGroupCount,
                      )}
                    </span>
                  </div>
                </div>

                {/* Quick Content Type Tabs */}
                <div
                  className={`${styles.templateActions} ${styles.templateActionsSpacingBottom}`}
                >
                  <SharePointButton
                    variant={
                      contentTypeFilter === ContentType.Alert
                        ? "primary"
                        : "secondary"
                    }
                    onClick={() => setContentTypeFilter(ContentType.Alert)}
                  >
                    {stripFilterEmoji(
                      Text.format(
                        strings.ManageAlertsFilterAlertsLabel,
                        contentTypeCounts.alert,
                      ),
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
                    {stripFilterEmoji(
                      Text.format(
                        strings.ManageAlertsFilterDraftsLabel,
                        contentTypeCounts.draft,
                      ),
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
                    {stripFilterEmoji(
                      Text.format(
                        strings.ManageAlertsFilterTemplatesLabel,
                        contentTypeCounts.template,
                      ),
                    )}
                  </SharePointButton>
                  <SharePointButton
                    variant={
                      contentTypeFilter === "all" ? "primary" : "secondary"
                    }
                    onClick={() => setContentTypeFilter("all")}
                  >
                    {stripFilterEmoji(
                      Text.format(
                        strings.ManageAlertsFilterAllLabel,
                        contentTypeCounts.all,
                      ),
                    )}
                  </SharePointButton>
                </div>

                {/* Search Bar - Always Visible */}
                <div className={styles.searchBar} ref={searchInputContainerRef}>
                  <SharePointInput
                    label={strings.ManageAlertsSearchLabel}
                    value={searchTerm}
                    onChange={setSearchTerm}
                    placeholder={stripFilterEmoji(
                      strings.ManageAlertsSearchPlaceholder,
                    )}
                    className={styles.searchInput}
                  />
                </div>

                {/* Collapsible Advanced Filters */}
                {showFilters && (
                  <div
                    id={advancedFiltersId}
                    className={styles.advancedFilters}
                  >
                    <div className={styles.filterGrid}>
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
                            label: stripFilterEmoji(
                              strings.ManageAlertsPriorityAllOption,
                            ),
                          },
                          {
                            value: AlertPriority.Critical,
                            label: strings.PriorityCritical,
                          },
                          {
                            value: AlertPriority.High,
                            label: strings.PriorityHigh,
                          },
                          {
                            value: AlertPriority.Medium,
                            label: strings.PriorityMedium,
                          },
                          {
                            value: AlertPriority.Low,
                            label: strings.PriorityLow,
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
                            label: stripFilterEmoji(
                              strings.ManageAlertsAlertTypeAllOption,
                            ),
                          },
                          ...alertTypes.map((type) => ({
                            value: type.name,
                            label: type.name,
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
                            label: stripFilterEmoji(
                              strings.ManageAlertsStatusAllOption,
                            ),
                          },
                          {
                            value: "active",
                            label: strings.StatusActive,
                          },
                          {
                            value: "expired",
                            label: strings.StatusExpired,
                          },
                          {
                            value: "scheduled",
                            label: strings.StatusScheduled,
                          },
                        ]}
                      />

                      {/* Content Status Filter */}
                      <SharePointSelect
                        label={strings.ManageAlertsApprovalStatusLabel}
                        value={contentStatusFilter}
                        onChange={(value) =>
                          setContentStatusFilter(value as ContentStatus | "all")
                        }
                        options={[
                          {
                            value: "all",
                            label: stripFilterEmoji(
                              strings.ManageAlertsApprovalAnyStatusOption,
                            ),
                          },
                          {
                            value: ContentStatus.Draft,
                            label: stripFilterEmoji(
                              strings.ManageAlertsApprovalDraftOption,
                            ),
                          },
                          {
                            value: ContentStatus.PendingReview,
                            label: stripFilterEmoji(
                              strings.ManageAlertsApprovalPendingOption,
                            ),
                          },
                          {
                            value: ContentStatus.Approved,
                            label: stripFilterEmoji(
                              strings.ManageAlertsApprovalApprovedOption,
                            ),
                          },
                          {
                            value: ContentStatus.Rejected,
                            label: stripFilterEmoji(
                              strings.ManageAlertsApprovalRejectedOption,
                            ),
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
                            label: stripFilterEmoji(
                              strings.ManageAlertsLanguageAllOption,
                            ),
                          },
                          {
                            value: TargetLanguage.All,
                            label: stripFilterEmoji(
                              strings.ManageAlertsLanguageMultiOption,
                            ),
                          },
                          ...supportedLanguages.map((lang) => ({
                            value: lang.code,
                            label: lang.nativeName,
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
                            label: stripFilterEmoji(
                              strings.ManageAlertsNotificationAllOption,
                            ),
                          },
                          {
                            value: NotificationType.None,
                            label: stripFilterEmoji(
                              strings.ManageAlertsNotificationNoneOption,
                            ),
                          },
                          {
                            value: NotificationType.Browser,
                            label: stripFilterEmoji(
                              strings.ManageAlertsNotificationBrowserOption,
                            ),
                          },
                          {
                            value: NotificationType.Email,
                            label: stripFilterEmoji(
                              strings.ManageAlertsNotificationEmailOption,
                            ),
                          },
                          {
                            value: NotificationType.Both,
                            label: stripFilterEmoji(
                              strings.ManageAlertsNotificationBothOption,
                            ),
                          },
                        ]}
                      />

                      {/* Date Filter */}
                      <SharePointSelect
                        label={strings.ManageAlertsCreatedDateFilterLabel}
                        value={dateFilter}
                        onChange={(value) =>
                          setDateFilter(
                            value as
                              | "all"
                              | "today"
                              | "week"
                              | "month"
                              | "custom",
                          )
                        }
                        options={[
                          {
                            value: "all",
                            label: stripFilterEmoji(
                              strings.ManageAlertsDateAllOption,
                            ),
                          },
                          {
                            value: "today",
                            label: stripFilterEmoji(
                              strings.ManageAlertsDateTodayOption,
                            ),
                          },
                          {
                            value: "week",
                            label: stripFilterEmoji(
                              strings.ManageAlertsDateWeekOption,
                            ),
                          },
                          {
                            value: "month",
                            label: stripFilterEmoji(
                              strings.ManageAlertsDateMonthOption,
                            ),
                          },
                          {
                            value: "custom",
                            label: stripFilterEmoji(
                              strings.ManageAlertsDateCustomOption,
                            ),
                          },
                        ]}
                      />
                    </div>

                    {/* Custom Date Range */}
                    {dateFilter === "custom" && (
                      <div className={styles.dateRangeFilters}>
                        <SharePointInput
                          label={strings.ManageAlertsFromDateLabel}
                          type="date"
                          value={customDateFrom}
                          onChange={setCustomDateFrom}
                        />
                        <SharePointInput
                          label={strings.ManageAlertsToDateLabel}
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
                        onClick={resetFilters}
                        disabled={!hasActiveFilters}
                      >
                        {strings.ManageAlertsClearFiltersButton}
                      </SharePointButton>
                    </div>
                  </div>
                )}
              </div>

              {/* Empty state for filtered drafts view */}
              {contentTypeFilter === ContentType.Draft &&
                groupedAlerts.length === 0 && (
                  <EmptyState
                    icon="Edit"
                    title={strings.ManageAlertsNoDraftsTitle}
                    description={strings.CreateAlertDraftsEmptyMessage}
                    actionText={strings.ManageAlertsEmptyCreateFirstButton}
                    onAction={() => setActiveTab("create")}
                  />
                )}

              {/* Empty state for filtered templates view */}
              {contentTypeFilter === ContentType.Template &&
                groupedAlerts.length === 0 && (
                  <EmptyState
                    icon="PageListSolid"
                    title={strings.ManageAlertsNoTemplatesTitle}
                    description={strings.AlertTemplatesEmptyDescription}
                    actionText={strings.ManageAlertsEmptyCreateFirstButton}
                    onAction={() => setActiveTab("create")}
                  />
                )}

              {/* Empty state when filters return no results */}
              {contentTypeFilter !== ContentType.Draft &&
                contentTypeFilter !== ContentType.Template &&
                groupedAlerts.length === 0 && (
                  <EmptyState
                    icon="FilterSolid"
                    title={strings.AlertTemplatesNoResultsTitle}
                    description={strings.AlertTemplatesNoResultsDescription}
                    actionText={strings.ManageAlertsClearFiltersButton}
                    onAction={resetFilters}
                  />
                )}

              {groupedAlerts.map((alert, index) => (
                <ManageAlertCard
                  key={alert.id}
                  alert={alert}
                  isSelected={selectedAlertIds.has(alert.id)}
                  isBusy={
                    isLoadingAlerts ||
                    isBulkActionInFlight ||
                    busyAlertIds.includes(alert.id)
                  }
                  onSelectionChange={(checked) => {
                    if (checked) {
                      setSelectedAlerts((prev) =>
                        prev.includes(alert.id) ? prev : [...prev, alert.id],
                      );
                    } else {
                      setSelectedAlerts((prev) =>
                        prev.filter((id) => id !== alert.id),
                      );
                    }
                  }}
                  onEdit={() => handleEditAlert(alert)}
                  onDelete={() => handleDeleteAlert(alert.id, alert.title)}
                  onPublishDraft={() => handlePublishDraft(alert)}
                  onSubmitForReview={() => handleSubmitForReview(alert)}
                  onApprove={() => handleApprove(alert)}
                  onReject={() => handleReject(alert)}
                  priorityLabel={
                    priorityLabelMap.get(String(alert.priority)) ||
                    alert.priority
                  }
                  alertTypeStyle={alertTypeStyleMap.get(alert.AlertType)}
                  supportedLanguageMap={supportedLanguageMap}
                  // Sort controls - only show when in manual sort mode
                  sortOrder={alert.sortOrder ?? index}
                  onSortOrderChange={manageSortMode === "manual" ? 
                    (newOrder) => handleSortOrderChange(alert.id, newOrder) : undefined}
                  onMoveToTop={manageSortMode === "manual" ? 
                    () => handleMoveToTop(alert.id) : undefined}
                  onMoveToBottom={manageSortMode === "manual" ? 
                    () => handleMoveToBottom(alert.id) : undefined}
                  isFirst={index === 0}
                  isLast={index === groupedAlerts.length - 1}
                  // Drag and drop - only in manual mode
                  isDragging={draggingAlertId === alert.id}
                  isDragOver={dragOverAlertId === alert.id}
                  onDragStart={manageSortMode === "manual" ? 
                    () => handleDragStart(alert.id) : undefined}
                  onDragEnd={handleDragEnd}
                  onDragOver={manageSortMode === "manual" ? 
                    () => handleDragOver(alert.id) : undefined}
                  onDrop={manageSortMode === "manual" ? 
                    () => handleDrop(alert.id) : undefined}
                />
              ))}
            </div>
          )}
        </div>
      )}

      {/* Edit Alert Form — rendered inline, replaces the parent dialog content */}
      {editingAlert && (
        <div className={styles.editAlertForm}>
          <AlertEditorForm
            mode="manage"
            alert={editingAlert}
            setAlert={setEditingAlertData}
            errors={editErrors}
            setErrors={setEditErrors}
            alertTypes={alertTypes}
            alertTypeOptions={alertTypeOptions}
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
            tenantDefaultLanguage={tenantDefaultLanguage}
            copilotEnabled={copilotEnabled}
            copilotService={copilotService}
            copilotAvailability={copilotAvailability}
            notificationService={notificationService}
            isBusy={isEditingAlert}
            showPreview={showPreview}
            setShowPreview={setShowPreview}
            imageFolderName={editingAlert.languageGroup || editingAlert.title}
            disableImageUpload={isCrossSiteAlert}
            renderMiddleSections={
              <AttachmentManager
                context={context}
                alertService={alertService}
                listId={editingAlertListId || alertsListId}
                itemId={editingItemId}
                siteId={editingAlertSiteId || undefined}
                attachments={editingAlert.attachments || []}
                onAttachmentsChange={(newAttachments) => {
                  setEditingAlertData((prev) => ({
                    ...prev,
                    attachments: newAttachments,
                  }));
                }}
              />
            }
            actionButtons={
              <>
                <SharePointButton
                  variant="primary"
                  onClick={handleSaveEdit}
                  disabled={isEditingAlert}
                  icon={<Save24Regular />}
                >
                  {isEditingAlert
                    ? strings.ManageAlertsSavingChangesLabel
                    : strings.ManageAlertsSaveChangesButton}
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={handleCancelEdit}
                  disabled={isEditingAlert}
                >
                  {strings.Cancel}
                </SharePointButton>
              </>
            }
            showScheduleSummary={true}
          />
        </div>
      )}
      {dialogs}
    </>
  );
};

export default ManageAlertsTab;
