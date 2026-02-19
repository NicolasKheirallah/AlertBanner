import * as React from "react";
import {
  Add24Regular,
  LocalLanguage24Regular,
  Wrench24Regular,
  Globe24Regular,
  ArrowClockwise24Regular,
  ArrowRight24Regular,
} from "@fluentui/react-icons";
import {
  Checkbox as FluentCheckbox,
  MessageBar as FluentMessageBar,
  MessageBarType,
} from "@fluentui/react";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
} from "../../UI/SharePointControls";
import {
  SharePointAlertService,
  IRepairResult,
} from "../../Services/SharePointAlertService";
import { StorageService } from "../../Services/StorageService";
import LanguageFieldManager from "../../UI/LanguageFieldManager";
import {
  LanguageAwarenessService,
  ISupportedLanguage,
} from "../../Services/LanguageAwarenessService";
import {
  DEFAULT_LANGUAGE_POLICY,
  ILanguagePolicy,
} from "../../Services/LanguagePolicyService";
import { NotificationService } from "../../Services/NotificationService";
import ProgressIndicator, {
  StepStatus,
  IProgressStep,
} from "../../UI/ProgressIndicator";
import RepairDialog from "../../UI/RepairDialog";
import PermissionStatus from "../../UI/PermissionStatus";
import { logger } from "../../Services/LoggerService";
import { EmailNotificationService } from "../../Services/EmailNotificationService";
import { Text as CoreText } from "@microsoft/sp-core-library";
import styles from "../AlertSettings.module.scss";
import { CAROUSEL_CONFIG } from "../../Utils/AppConstants";
import {
  AlertPriority,
  TranslationStatus,
  IPriorityColorConfig,
} from "../../Alerts/IAlerts";
import ColorPicker from "../../UI/ColorPicker";
import {
  useAlertsState,
  useAlertsDispatch,
  AlertSortMode,
} from "../../Context/AlertsContext";
import * as strings from "AlertBannerApplicationCustomizerStrings";

const cardStyles = {
  languageGrid: styles.f2LanguageGrid,
  languageItem: styles.f2LanguageItem,
  languageInfo: styles.f2LanguageInfo,
  languageName: styles.f2LanguageName,
  languageCode: styles.f2LanguageCode,
  cardHeader: styles.f2CardHeaderInline,
  cardContent: styles.f2CardContent,
  hintText: styles.f2HintText,
};

const cx = (...classes: Array<string | undefined | false>): string =>
  classes.filter(Boolean).join(" ");

const Card: React.FC<{ children?: React.ReactNode; className?: string }> = ({
  children,
  className,
}) => <div className={cx(styles.f2Card, className)}>{children}</div>;

const CardHeader: React.FC<{
  image?: React.ReactNode;
  header?: React.ReactNode;
  description?: React.ReactNode;
}> = ({ image, header, description }) => (
  <div className={styles.f2CardHeader}>
    {image && <div>{image}</div>}
    <div className={styles.f2CardHeaderText}>
      {header}
      {description}
    </div>
  </div>
);

const CardPreview: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div>{children}</div>;

const Text: React.FC<{
  children?: React.ReactNode;
  size?: number;
  weight?: "regular" | "medium" | "semibold" | "bold";
  className?: string;
}> = ({ children, size, weight, className }) => (
  <span
    className={cx(
      styles.f2Text,
      size === 100 && styles.f2Text100,
      size === 200 && styles.f2Text200,
      size === 300 && styles.f2Text300,
      size === 400 && styles.f2Text400,
      size === 500 && styles.f2Text500,
      weight === "regular" && styles.f2TextRegular,
      weight === "medium" && styles.f2TextMedium,
      weight === "semibold" && styles.f2TextSemibold,
      weight === "bold" && styles.f2TextBold,
      className,
    )}
  >
    {children}
  </span>
);

const Checkbox: React.FC<{
  checked?: boolean;
  label?: React.ReactNode;
  disabled?: boolean;
  onChange?: (
    event: React.FormEvent<HTMLElement> | undefined,
    data: { checked?: boolean },
  ) => void;
}> = ({ checked, label, disabled, onChange }) => (
  <FluentCheckbox
    checked={checked}
    disabled={disabled}
    label={typeof label === "string" ? label : undefined}
    onRenderLabel={
      typeof label === "string" || typeof label === "undefined"
        ? undefined
        : () => <>{label}</>
    }
    onChange={(event, isChecked) =>
      onChange?.(event as React.FormEvent<HTMLElement>, { checked: isChecked })
    }
  />
);

const MessageBar: React.FC<{
  intent?: "error" | "warning" | "success" | "info";
  children?: React.ReactNode;
  ["aria-live"]?: "off" | "polite" | "assertive";
}> = ({ intent = "info", children, ...rest }) => (
  <FluentMessageBar
    messageBarType={
      intent === "error"
        ? MessageBarType.error
        : intent === "warning"
          ? MessageBarType.warning
          : intent === "success"
            ? MessageBarType.success
            : MessageBarType.info
    }
    isMultiline
    {...rest}
  >
    {children}
  </FluentMessageBar>
);

const MessageBarBody: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <>{children}</>;
const MessageBarTitle: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div className={styles.f2MessageTitle}>{children}</div>;

export interface ISettingsData {
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  emailServiceAccount?: string;
  copilotEnabled?: boolean;
}

export interface ISettingsTabProps {
  settings: ISettingsData;
  setSettings: React.Dispatch<React.SetStateAction<ISettingsData>>;
  alertsListExists: boolean | null;
  setAlertsListExists: React.Dispatch<React.SetStateAction<boolean | null>>;
  alertTypesListExists: boolean | null;
  setAlertTypesListExists: React.Dispatch<React.SetStateAction<boolean | null>>;
  isCheckingLists: boolean;
  setIsCheckingLists: React.Dispatch<React.SetStateAction<boolean>>;
  isCreatingLists: boolean;
  setIsCreatingLists: React.Dispatch<React.SetStateAction<boolean>>;
  alertService: SharePointAlertService;
  onSettingsChange: (settings: ISettingsData) => void;
  onLanguageChange?: (languages: string[]) => void;
  onDirtyStateChange?: (isDirty: boolean) => void;
  canEdit?: boolean;
  context?: any; // ApplicationCustomizerContext for notifications
}

const SettingsTab: React.FC<ISettingsTabProps> = ({
  settings,
  setSettings,
  alertsListExists,
  setAlertsListExists,
  alertTypesListExists,
  setAlertTypesListExists,
  isCheckingLists,
  setIsCheckingLists,
  isCreatingLists,
  setIsCreatingLists,
  alertService,
  onSettingsChange,
  onLanguageChange,
  onDirtyStateChange,
  canEdit = true,
  context,
}) => {
  const alertsState = useAlertsState();
  const { updateCarouselSettings, updatePriorityBorderColors, updateSortMode } =
    useAlertsDispatch();
  const storageService = React.useRef<StorageService>(
    StorageService.getInstance(),
  );
  const [draftSettings, setDraftSettings] =
    React.useState<ISettingsData>(settings);
  const [carouselEnabled, setCarouselEnabled] = React.useState(false);
  const [carouselInterval, setCarouselInterval] = React.useState("5");
  const [savedCarouselSettings, setSavedCarouselSettings] = React.useState<{
    enabled: boolean;
    interval: number;
  }>({ enabled: false, interval: 5 });
  const [isRepairDialogOpen, setIsRepairDialogOpen] = React.useState(false);
  const [preCreationLanguages, setPreCreationLanguages] = React.useState<
    string[]
  >(["en-us"]); // English selected by default
  const [creationSteps, setCreationSteps] = React.useState<IProgressStep[]>([]);
  const [languagePolicy, setLanguagePolicy] = React.useState<ILanguagePolicy>(
    DEFAULT_LANGUAGE_POLICY,
  );
  const [savedLanguagePolicy, setSavedLanguagePolicy] =
    React.useState<ILanguagePolicy>(DEFAULT_LANGUAGE_POLICY);
  const [policyLanguages, setPolicyLanguages] = React.useState<
    ISupportedLanguage[]
  >([]);
  const [isSavingPolicy, setIsSavingPolicy] = React.useState(false);
  const [isSavingAll, setIsSavingAll] = React.useState(false);
  const [showAdvancedLocalization, setShowAdvancedLocalization] =
    React.useState(false);
  const [showAdvancedMaintenance, setShowAdvancedMaintenance] =
    React.useState(false);
  const [lastSavedAt, setLastSavedAt] = React.useState<Date | null>(null);
  const [priorityBorderColors, setPriorityBorderColors] = React.useState<
    Record<AlertPriority, IPriorityColorConfig>
  >({
    [AlertPriority.Critical]: { borderColor: "#d13438" },
    [AlertPriority.High]: { borderColor: "#f7630c" },
    [AlertPriority.Medium]: { borderColor: "#0078d4" },
    [AlertPriority.Low]: { borderColor: "#107c10" },
  });
  const [savedPriorityBorderColors, setSavedPriorityBorderColors] =
    React.useState<Record<AlertPriority, IPriorityColorConfig>>({
      [AlertPriority.Critical]: { borderColor: "#d13438" },
      [AlertPriority.High]: { borderColor: "#f7630c" },
      [AlertPriority.Medium]: { borderColor: "#0078d4" },
      [AlertPriority.Low]: { borderColor: "#107c10" },
    });
  const importInputRef = React.useRef<HTMLInputElement | null>(null);
  const isSiteAdmin = !!context?.pageContext?.legacyPageContext?.isSiteAdmin;
  const canPersistSettings = canEdit && isSiteAdmin;
  const notificationService = React.useMemo(
    () => (context ? NotificationService.getInstance(context) : null),
    [context],
  );

  React.useEffect(() => {
    setDraftSettings(settings);
  }, [settings]);

  // Load carousel settings from StorageService on mount
  React.useEffect(() => {
    const savedCarouselEnabled =
      storageService.current.getFromLocalStorage<boolean>("carouselEnabled");
    const savedCarouselInterval =
      storageService.current.getFromLocalStorage<number>("carouselInterval");

    if (savedCarouselEnabled !== null) {
      setCarouselEnabled(savedCarouselEnabled);
      setSavedCarouselSettings((prev) => ({
        ...prev,
        enabled: savedCarouselEnabled,
      }));
    }
    if (
      savedCarouselInterval &&
      savedCarouselInterval >= CAROUSEL_CONFIG.MIN_INTERVAL &&
      savedCarouselInterval <= CAROUSEL_CONFIG.MAX_INTERVAL
    ) {
      const seconds = savedCarouselInterval / 1000;
      setCarouselInterval(seconds.toString());
      setSavedCarouselSettings((prev) => ({ ...prev, interval: seconds }));
    }

    // Load priority border colors
    const savedPriorityColors = storageService.current.getFromLocalStorage<
      Record<AlertPriority, IPriorityColorConfig>
    >("priorityBorderColors");
    if (savedPriorityColors) {
      setPriorityBorderColors(savedPriorityColors);
      setSavedPriorityBorderColors(savedPriorityColors);
    }
  }, []);

  React.useEffect(() => {
    if (!alertsListExists) return;
    let isMounted = true;
    const loadPolicy = async () => {
      try {
        const [policy, supportedLanguageCodes] = await Promise.all([
          alertService.getLanguagePolicy(),
          alertService.getSupportedLanguages(),
        ]);

        if (!isMounted) return;
        setLanguagePolicy(policy);
        setSavedLanguagePolicy(policy);

        const baseLanguages = LanguageAwarenessService.getSupportedLanguages();
        const updatedLanguages = baseLanguages.map((lang) => ({
          ...lang,
          columnExists:
            supportedLanguageCodes.includes(lang.code) || lang.code === "en-us",
          isSupported:
            supportedLanguageCodes.includes(lang.code) || lang.code === "en-us",
        }));
        setPolicyLanguages(
          updatedLanguages.filter(
            (lang) => lang.isSupported && lang.columnExists,
          ),
        );
      } catch (error) {
        logger.warn("SettingsTab", "Failed to load language policy", error);
      }
    };

    loadPolicy();
    return () => {
      isMounted = false;
    };
  }, [alertService, alertsListExists]);

  const handleCarouselEnabledChange = React.useCallback((checked: boolean) => {
    setCarouselEnabled(checked);
  }, []);

  const handleCarouselIntervalChange = React.useCallback((value: string) => {
    setCarouselInterval(value);
  }, []);

  const handleSettingsChange = React.useCallback(
    (newSettings: Partial<ISettingsData>) => {
      setDraftSettings((prev) => ({ ...prev, ...newSettings }));
    },
    [],
  );

  const isValidEmailAddress = React.useCallback((value: string): boolean => {
    if (!value.trim()) {
      return true;
    }
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim());
  }, []);

  const carouselIntervalNumber = React.useMemo(
    () => Number.parseInt(carouselInterval, 10),
    [carouselInterval],
  );

  const preflightErrors = React.useMemo(() => {
    const errors: string[] = [];

    if (!Number.isInteger(carouselIntervalNumber)) {
      errors.push(strings.SettingsValidationCarouselNumber);
    } else if (
      carouselIntervalNumber < CAROUSEL_CONFIG.MIN_INTERVAL / 1000 ||
      carouselIntervalNumber > CAROUSEL_CONFIG.MAX_INTERVAL / 1000
    ) {
      errors.push(
        strings.SettingsValidationCarouselRange.replace(
          "{0}",
          String(CAROUSEL_CONFIG.MIN_INTERVAL / 1000),
        ).replace("{1}", String(CAROUSEL_CONFIG.MAX_INTERVAL / 1000)),
      );
    }

    if (
      draftSettings.notificationsEnabled &&
      !isValidEmailAddress(draftSettings.emailServiceAccount || "")
    ) {
      errors.push(strings.SettingsValidationEmail);
    }

    try {
      JSON.parse(draftSettings.alertTypesJson || "[]");
    } catch (error) {
      errors.push(strings.SettingsValidationAlertTypesJson);
    }

    return errors;
  }, [
    carouselIntervalNumber,
    draftSettings.alertTypesJson,
    draftSettings.emailServiceAccount,
    draftSettings.notificationsEnabled,
    isValidEmailAddress,
  ]);

  const preflightWarnings = React.useMemo(() => {
    const warnings: string[] = [];

    if (
      draftSettings.notificationsEnabled &&
      !(draftSettings.emailServiceAccount || "").trim()
    ) {
      warnings.push(strings.SettingsWarningNotificationsWithoutEmail);
    }

    if (!canPersistSettings) {
      warnings.push(strings.SettingsReadOnlyWarning);
    }

    return warnings;
  }, [
    canPersistSettings,
    draftSettings.emailServiceAccount,
    draftSettings.notificationsEnabled,
  ]);

  const settingsDirty = React.useMemo(() => {
    return JSON.stringify(draftSettings) !== JSON.stringify(settings);
  }, [draftSettings, settings]);

  const carouselDirty = React.useMemo(() => {
    return (
      carouselEnabled !== savedCarouselSettings.enabled ||
      carouselIntervalNumber !== savedCarouselSettings.interval
    );
  }, [
    carouselEnabled,
    carouselIntervalNumber,
    savedCarouselSettings.enabled,
    savedCarouselSettings.interval,
  ]);

  const languagePolicyDirty = React.useMemo(() => {
    return (
      JSON.stringify(languagePolicy || {}) !==
      JSON.stringify(savedLanguagePolicy || {})
    );
  }, [languagePolicy, savedLanguagePolicy]);

  const priorityColorsDirty = React.useMemo(() => {
    return (
      JSON.stringify(priorityBorderColors) !==
      JSON.stringify(savedPriorityBorderColors)
    );
  }, [priorityBorderColors, savedPriorityBorderColors]);

  const hasUnsavedChanges =
    settingsDirty ||
    carouselDirty ||
    languagePolicyDirty ||
    priorityColorsDirty;

  React.useEffect(() => {
    onDirtyStateChange?.(hasUnsavedChanges);
  }, [hasUnsavedChanges, onDirtyStateChange]);

  const completenessOptions = React.useMemo(
    () => [
      {
        value: "allSelectedComplete",
        label: strings.LanguagePolicyCompletenessAll,
      },
      {
        value: "atLeastOneComplete",
        label: strings.LanguagePolicyCompletenessAtLeastOne,
      },
      {
        value: "requireDefaultLanguageComplete",
        label: strings.LanguagePolicyCompletenessDefault,
      },
    ],
    [],
  );

  const fallbackOptions = React.useMemo(() => {
    const options = [
      {
        value: "tenant-default",
        label: strings.LanguagePolicyFallbackTenantDefault,
      },
    ];
    policyLanguages.forEach((lang) => {
      options.push({
        value: lang.code,
        label: `${lang.flag} ${lang.nativeName} (${lang.name})`,
      });
    });
    return options;
  }, [policyLanguages]);

  const workflowStatusOptions = React.useMemo(
    () => [
      { value: TranslationStatus.Draft, label: strings.TranslationStatusDraft },
      {
        value: TranslationStatus.InReview,
        label: strings.TranslationStatusInReview,
      },
      {
        value: TranslationStatus.Approved,
        label: strings.TranslationStatusApproved,
      },
    ],
    [],
  );

  const handleSaveLanguagePolicy = React.useCallback(async () => {
    setIsSavingPolicy(true);
    try {
      await alertService.saveLanguagePolicy(languagePolicy);
      setSavedLanguagePolicy(languagePolicy);
      setLastSavedAt(new Date());
      notificationService?.showSuccess(
        strings.LanguagePolicySavedSuccess,
        strings.LanguagePolicyTitle,
      );
    } catch (error) {
      logger.error("SettingsTab", "Failed to save language policy", error);
      notificationService?.showError(
        strings.LanguagePolicySavedFailed,
        strings.LanguagePolicyTitle,
      );
    } finally {
      setIsSavingPolicy(false);
    }
  }, [alertService, languagePolicy, notificationService]);

  const syncCarouselToContext = React.useCallback(
    (enabled: boolean, intervalSeconds: number) => {
      storageService.current.saveToLocalStorage("carouselEnabled", enabled);
      storageService.current.saveToLocalStorage(
        "carouselInterval",
        intervalSeconds * 1000,
      );
      updateCarouselSettings({
        carouselEnabled: enabled,
        carouselInterval: intervalSeconds * 1000,
      });
      setSavedCarouselSettings({
        enabled,
        interval: intervalSeconds,
      });
    },
    [updateCarouselSettings],
  );

  const handleDiscardChanges = React.useCallback(() => {
    setDraftSettings(settings);
    setCarouselEnabled(savedCarouselSettings.enabled);
    setCarouselInterval(String(savedCarouselSettings.interval));
    setLanguagePolicy(savedLanguagePolicy);
    setPriorityBorderColors(savedPriorityBorderColors);
  }, [
    savedCarouselSettings,
    savedLanguagePolicy,
    savedPriorityBorderColors,
    settings,
  ]);

  const handleResetToDefaults = React.useCallback(() => {
    setDraftSettings((prev) => ({
      ...prev,
      userTargetingEnabled: true,
      notificationsEnabled: false,
      enableTargetSite: false,
      emailServiceAccount: "",
      copilotEnabled: false,
    }));
    setCarouselEnabled(false);
    setCarouselInterval("5");
    setLanguagePolicy(DEFAULT_LANGUAGE_POLICY);
  }, []);

  const handleSaveAll = React.useCallback(async () => {
    if (preflightErrors.length > 0 || !canPersistSettings) {
      return;
    }

    setIsSavingAll(true);
    try {
      if (settingsDirty) {
        setSettings(draftSettings);
        onSettingsChange(draftSettings);
      }

      if (carouselDirty) {
        syncCarouselToContext(carouselEnabled, carouselIntervalNumber);
      }

      if (languagePolicyDirty) {
        await alertService.saveLanguagePolicy(languagePolicy);
        setSavedLanguagePolicy(languagePolicy);
      }

      if (priorityColorsDirty) {
        storageService.current.saveToLocalStorage(
          "priorityBorderColors",
          priorityBorderColors,
        );
        setSavedPriorityBorderColors(priorityBorderColors);
        updatePriorityBorderColors(priorityBorderColors);
      }

      setLastSavedAt(new Date());
      notificationService?.showSuccess(
        strings.SettingsSavedSuccess,
        strings.SettingsTabTitle,
      );
    } catch (error) {
      logger.error("SettingsTab", "Failed to save settings", error);
      notificationService?.showError(
        strings.FailedToSaveSettings,
        strings.SettingsTabTitle,
      );
    } finally {
      setIsSavingAll(false);
    }
  }, [
    alertService,
    canPersistSettings,
    carouselDirty,
    carouselEnabled,
    carouselIntervalNumber,
    draftSettings,
    languagePolicy,
    languagePolicyDirty,
    notificationService,
    onSettingsChange,
    preflightErrors.length,
    setSettings,
    settingsDirty,
    syncCarouselToContext,
  ]);

  const handleExportSettings = React.useCallback(() => {
    const payload = {
      schemaVersion: 2,
      exportedAt: new Date().toISOString(),
      settings: draftSettings,
      carousel: {
        enabled: carouselEnabled,
        intervalSeconds: Number.isInteger(carouselIntervalNumber)
          ? carouselIntervalNumber
          : 5,
      },
      languagePolicy,
      priorityBorderColors,
    };

    try {
      const blob = new Blob([JSON.stringify(payload, null, 2)], {
        type: "application/json",
      });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "alert-banner-settings.json";
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
      URL.revokeObjectURL(url);
      notificationService?.showSuccess(
        strings.SettingsExportSuccess,
        strings.SettingsTabTitle,
      );
    } catch (error) {
      logger.error("SettingsTab", "Failed to export settings", error);
      notificationService?.showError(
        strings.SettingsExportFailed,
        strings.SettingsTabTitle,
      );
    }
  }, [
    carouselEnabled,
    carouselIntervalNumber,
    draftSettings,
    languagePolicy,
    notificationService,
  ]);

  const handleImportSettingsClick = React.useCallback(() => {
    importInputRef.current?.click();
  }, []);

  const handleImportSettingsFile = React.useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) {
        return;
      }

      try {
        const fileText = await file.text();
        const parsed = JSON.parse(fileText) as {
          settings?: Partial<ISettingsData>;
          carousel?: { enabled?: boolean; intervalSeconds?: number };
          languagePolicy?: ILanguagePolicy;
          priorityBorderColors?: Record<AlertPriority, IPriorityColorConfig>;
        };

        if (parsed.settings) {
          setDraftSettings((prev) => ({ ...prev, ...parsed.settings }));
        }
        if (parsed.carousel) {
          if (typeof parsed.carousel.enabled === "boolean") {
            setCarouselEnabled(parsed.carousel.enabled);
          }
          if (
            typeof parsed.carousel.intervalSeconds === "number" &&
            Number.isInteger(parsed.carousel.intervalSeconds)
          ) {
            setCarouselInterval(String(parsed.carousel.intervalSeconds));
          }
        }
        if (parsed.languagePolicy) {
          setLanguagePolicy({
            ...DEFAULT_LANGUAGE_POLICY,
            ...parsed.languagePolicy,
          });
        }
        if (parsed.priorityBorderColors) {
          setPriorityBorderColors(parsed.priorityBorderColors);
        }

        notificationService?.showSuccess(
          strings.SettingsImportSuccess,
          strings.SettingsTabTitle,
        );
      } catch (error) {
        logger.error("SettingsTab", "Failed to import settings", error);
        notificationService?.showError(
          strings.SettingsImportFailed,
          strings.SettingsTabTitle,
        );
      } finally {
        event.target.value = "";
      }
    },
    [notificationService],
  );

  const checkListsExistence = React.useCallback(async () => {
    setIsCheckingLists(true);
    try {
      // Use the new detailed check method
      const listStatus = await alertService.checkListsNeeded();
      const currentSite = listStatus[0]; // Should be current site

      if (currentSite) {
        setAlertsListExists(currentSite.needsAlerts ? false : true);
        setAlertTypesListExists(currentSite.needsTypes ? false : true);
      } else {
        // Fallback to old method
        const [alertsTest, typesTest] = await Promise.allSettled([
          alertService.getAlerts(),
          alertService.getAlertTypes(),
        ]);

        setAlertsListExists(alertsTest.status === "fulfilled");
        setAlertTypesListExists(typesTest.status === "fulfilled");
      }
    } catch (error) {
      logger.error("SettingsTab", "Error checking lists", error);
      // Fallback: assume lists don't exist if there's an error
      setAlertsListExists(false);
      setAlertTypesListExists(false);
    } finally {
      setIsCheckingLists(false);
    }
  }, [
    alertService,
    setAlertsListExists,
    setAlertTypesListExists,
    setIsCheckingLists,
  ]);

  const handleCreateLists = React.useCallback(async () => {
    setIsCreatingLists(true);

    // Initialize progress steps
    const steps: IProgressStep[] = [
      {
        id: "check-lists",
        name: strings.SettingsProgressCheckListsName,
        description: strings.SettingsProgressCheckListsDescription,
        status: StepStatus.InProgress,
      },
      {
        id: "create-lists",
        name: strings.SettingsProgressCreateListsName,
        description: strings.SettingsProgressCreateListsDescription,
        status: StepStatus.Pending,
      },
    ];

    // Add language steps if multiple languages selected
    if (preCreationLanguages.length > 1) {
      preCreationLanguages.forEach((lang) => {
        if (lang !== "en-us") {
          steps.push({
            id: `add-language-${lang}`,
            name: strings.SettingsProgressAddLanguageName.replace(
              "{0}",
              lang.toUpperCase(),
            ),
            description: strings.SettingsProgressAddLanguageDescription.replace(
              "{0}",
              lang,
            ),
            status: StepStatus.Pending,
          });
        }
      });
    }

    steps.push({
      id: "finalize",
      name: strings.SettingsProgressFinalizeName,
      description: strings.SettingsProgressFinalizeDescription,
      status: StepStatus.Pending,
    });

    setCreationSteps(steps);

    try {
      // First check what's needed
      const listStatus = await alertService.checkListsNeeded();
      const currentSite = listStatus[0];

      // Update first step as completed
      setCreationSteps((prev) =>
        prev.map((step) =>
          step.id === "check-lists"
            ? { ...step, status: StepStatus.Completed }
            : step,
        ),
      );

      if (
        !currentSite ||
        (!currentSite.needsAlerts && !currentSite.needsTypes)
      ) {
        if (notificationService) {
          notificationService.showInfo(
            strings.SettingsListsAlreadyExistMessage,
            strings.SettingsListsAlreadyExistTitle,
          );
        } else {
          logger.info("SettingsTab", strings.SettingsListsAlreadyExistMessage);
        }
        return;
      }

      // Start creating lists step
      setCreationSteps((prev) =>
        prev.map((step) =>
          step.id === "create-lists"
            ? { ...step, status: StepStatus.InProgress }
            : step,
        ),
      );

      // Initialize lists using the existing service method
      await alertService.initializeLists();

      // Complete create lists step
      setCreationSteps((prev) =>
        prev.map((step) =>
          step.id === "create-lists"
            ? { ...step, status: StepStatus.Completed }
            : step,
        ),
      );

      // Add selected language columns to the newly created lists
      if (
        preCreationLanguages.length > 1 ||
        !preCreationLanguages.includes("en-us")
      ) {
        for (const languageCode of preCreationLanguages) {
          if (languageCode !== "en-us") {
            // English is already included by default
            // Start language step
            setCreationSteps((prev) =>
              prev.map((step) =>
                step.id === `add-language-${languageCode}`
                  ? { ...step, status: StepStatus.InProgress }
                  : step,
              ),
            );

            try {
              await alertService.addLanguageSupport(languageCode);
              logger.debug(
                "SettingsTab",
                `Added ${languageCode} language columns during list creation`,
              );

              // Complete language step
              setCreationSteps((prev) =>
                prev.map((step) =>
                  step.id === `add-language-${languageCode}`
                    ? { ...step, status: StepStatus.Completed }
                    : step,
                ),
              );
            } catch (error) {
              logger.warn(
                "SettingsTab",
                `Failed to add ${languageCode} language columns`,
                error,
              );

              // Mark language step as failed
              setCreationSteps((prev) =>
                prev.map((step) =>
                  step.id === `add-language-${languageCode}`
                    ? {
                        ...step,
                        status: StepStatus.Failed,
                        error: error.message,
                      }
                    : step,
                ),
              );
            }
          }
        }
      }

      // Start finalize step
      setCreationSteps((prev) =>
        prev.map((step) =>
          step.id === "finalize"
            ? { ...step, status: StepStatus.InProgress }
            : step,
        ),
      );

      // Re-check lists after creation
      await checkListsExistence();

      // Complete finalize step
      setCreationSteps((prev) =>
        prev.map((step) =>
          step.id === "finalize"
            ? { ...step, status: StepStatus.Completed }
            : step,
        ),
      );

      // Success message
      const createdLists = [];
      if (currentSite.needsAlerts) createdLists.push("Alerts");
      if (currentSite.needsTypes && currentSite.isHomeSite)
        createdLists.push("AlertBannerTypes (Home Site only)");

      if (createdLists.length > 0) {
        const languageMessage =
          preCreationLanguages.length > 1
            ? ` with support for ${preCreationLanguages.length} languages (${preCreationLanguages.join(", ")})`
            : "";
        const successMessage = `Successfully created ${createdLists.join(" and ")} list${createdLists.length > 1 ? "s" : ""}${languageMessage} on this site.`;

        if (notificationService) {
          notificationService.showSuccess(
            successMessage,
            strings.SettingsListsCreatedTitle,
          );
        } else {
          logger.info("SettingsTab", successMessage);
        }
      }

      // Trigger language change callback to refresh other components
      if (onLanguageChange) {
        onLanguageChange(preCreationLanguages);
      }

      // Show informational message about AlertBannerTypes if not on home site
      if (!currentSite.isHomeSite && currentSite.needsAlerts) {
        const infoMessage =
          "Alerts list created successfully. Note: AlertBannerTypes list is only created on the SharePoint home site to maintain consistency across the tenant.";

        if (notificationService) {
          notificationService.showInfo(
            infoMessage,
            strings.SettingsAdditionalInformationTitle,
          );
        } else {
          logger.info("SettingsTab", infoMessage);
        }
      }
    } catch (error) {
      logger.error("SettingsTab", "Error creating lists", error);
      const errorMsg = error.message || error.toString();

      if (errorMsg.includes("PERMISSION_DENIED")) {
        const permissionError =
          "Permission denied: You need site owner or full control permissions to create SharePoint lists.";

        if (notificationService) {
          notificationService.showError(permissionError, strings.Error, [
            {
              text: strings.SettingsContactAdministrator,
              onClick: () => {
                window.open(
                  "mailto:?subject=SharePoint Permissions Required&body=I need permissions to create SharePoint lists for the Alert Banner system.",
                );
              },
            },
          ]);
        } else {
          logger.error("SettingsTab", permissionError);
        }
      } else {
        const generalError = `Failed to create some lists: ${errorMsg}`;

        if (notificationService) {
          notificationService.showError(generalError, strings.Error, [
            {
              text: strings.CreateAlertCreationRetryButton,
              onClick: () => handleCreateLists(),
            },
          ]);
        } else {
          logger.error("SettingsTab", generalError);
        }
      }
    } finally {
      setIsCreatingLists(false);
    }
  }, [alertService, checkListsExistence, setIsCreatingLists]);

  const handleOpenRepairDialog = React.useCallback(() => {
    setIsRepairDialogOpen(true);
  }, []);

  const handleCloseRepairDialog = React.useCallback(() => {
    setIsRepairDialogOpen(false);
  }, []);

  const handleRepairComplete = React.useCallback(
    async (result: IRepairResult) => {
      // Show appropriate notification based on result
      if (notificationService) {
        if (result.success) {
          if (result.details.warnings.length > 0) {
            notificationService.showWarning(
              result.message,
              strings.RepairDialogDialogTitleIssues,
            );
          } else {
            notificationService.showSuccess(
              result.message,
              strings.RepairDialogDialogTitleSuccess,
            );
          }
        } else {
          notificationService.showError(
            result.message,
            strings.RepairDialogResultFailureTitle,
          );
        }
      }

      // Re-check lists after repair to refresh the UI
      try {
        await checkListsExistence();
      } catch (error) {
        logger.warn(
          "SettingsTab",
          "Failed to refresh list status after repair",
          error,
        );
      }
    },
    [checkListsExistence, notificationService],
  );

  // Check lists on mount
  React.useEffect(() => {
    checkListsExistence();
  }, [checkListsExistence]);

  const scrollToSection = React.useCallback((sectionId: string) => {
    if (typeof document === "undefined") {
      return;
    }
    const section = document.getElementById(sectionId);
    section?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, []);

  return (
    <div className={styles.tabPane}>
      <div className={styles.settingsContent}>
        <div className={styles.settingsQuickNav} role="navigation">
          <SharePointButton
            variant="secondary"
            onClick={() => scrollToSection("settings-section-audience")}
          >
            {strings.SettingsSectionAudienceTitle}
          </SharePointButton>
          <SharePointButton
            variant="secondary"
            onClick={() => scrollToSection("settings-section-integrations")}
          >
            {strings.SettingsSectionIntegrationsTitle}
          </SharePointButton>
          <SharePointButton
            variant="secondary"
            onClick={() => scrollToSection("settings-section-localization")}
          >
            {strings.SettingsSectionLocalizationTitle}
          </SharePointButton>
          <SharePointButton
            variant="secondary"
            onClick={() => scrollToSection("settings-section-setup")}
          >
            {strings.SettingsSectionSetupTitle}
          </SharePointButton>
          <SharePointButton
            variant="secondary"
            onClick={() => scrollToSection("settings-section-maintenance")}
          >
            {strings.SettingsSectionMaintenanceTitle}
          </SharePointButton>
        </div>

        <div className={styles.settingsActionBar}>
          <div className={styles.settingsActionPrimary}>
            <SharePointButton
              variant="primary"
              onClick={handleSaveAll}
              disabled={
                !hasUnsavedChanges ||
                isSavingAll ||
                !canPersistSettings ||
                preflightErrors.length > 0
              }
            >
              {isSavingAll ? strings.SettingsSaving : strings.SettingsSaveAll}
            </SharePointButton>
            <SharePointButton
              variant="secondary"
              onClick={handleDiscardChanges}
              disabled={!hasUnsavedChanges || isSavingAll}
            >
              {strings.SettingsDiscardChanges}
            </SharePointButton>
            <SharePointButton
              variant="secondary"
              icon={<ArrowClockwise24Regular />}
              onClick={checkListsExistence}
              disabled={isCheckingLists}
            >
              {strings.Refresh}
            </SharePointButton>
          </div>
          <div className={styles.settingsActionMeta}>
            {hasUnsavedChanges && (
              <span className={styles.settingsUnsavedBadge}>
                {strings.SettingsUnsavedChangesLabel}
              </span>
            )}
            {lastSavedAt && (
              <span className={styles.settingsLastSaved}>
                {strings.SettingsLastSavedLabel.replace(
                  "{0}",
                  lastSavedAt.toLocaleTimeString(),
                )}
              </span>
            )}
          </div>
        </div>

        {preflightErrors.length > 0 && (
          <MessageBar intent="error" aria-live="polite">
            <MessageBarBody>
              <MessageBarTitle>
                {strings.SettingsPreflightErrorsTitle}
              </MessageBarTitle>
              <ul className={styles.settingsMessageList}>
                {preflightErrors.map((error) => (
                  <li key={error}>{error}</li>
                ))}
              </ul>
            </MessageBarBody>
          </MessageBar>
        )}

        {preflightWarnings.length > 0 && (
          <MessageBar intent="warning" aria-live="polite">
            <MessageBarBody>
              <MessageBarTitle>
                {strings.SettingsPreflightWarningsTitle}
              </MessageBarTitle>
              <ul className={styles.settingsMessageList}>
                {preflightWarnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </MessageBarBody>
          </MessageBar>
        )}

        {context && (
          <div id="settings-section-permissions">
            <SharePointSection title={strings.SettingsSectionPermissionsTitle}>
              <PermissionStatus context={context} />
            </SharePointSection>
          </div>
        )}

        <div id="settings-section-audience">
          <SharePointSection title={strings.SettingsSectionAudienceTitle}>
            <div className={styles.settingsGrid}>
              <SharePointToggle
                label={strings.EnableUserTargeting}
                checked={draftSettings.userTargetingEnabled}
                onChange={(checked) =>
                  handleSettingsChange({ userTargetingEnabled: checked })
                }
                description={strings.EnableUserTargetingDescription}
                disabled={!canPersistSettings}
              />

              <SharePointToggle
                label={strings.SettingsEnableTargetSitesLabel}
                checked={draftSettings.enableTargetSite}
                onChange={(checked) =>
                  handleSettingsChange({ enableTargetSite: checked })
                }
                description={strings.SettingsEnableTargetSitesDescription}
                disabled={!canPersistSettings}
              />
            </div>
          </SharePointSection>
        </div>

        <div id="settings-section-integrations">
          <SharePointSection title={strings.SettingsSectionIntegrationsTitle}>
            <div className={styles.settingsGrid}>
              <SharePointToggle
                label={strings.EnableNotifications}
                checked={draftSettings.notificationsEnabled}
                onChange={(checked) =>
                  handleSettingsChange({ notificationsEnabled: checked })
                }
                description={strings.SettingsEnableNotificationsDescription}
                disabled={!canPersistSettings}
              />

              {draftSettings.notificationsEnabled && (
                <div className={styles.settingsInlineStack}>
                  <SharePointInput
                    label={strings.SettingsEmailServiceAccountLabel}
                    value={draftSettings.emailServiceAccount || ""}
                    onChange={(value) =>
                      handleSettingsChange({ emailServiceAccount: value })
                    }
                    placeholder={strings.SettingsEmailServiceAccountPlaceholder}
                    description={strings.SettingsEmailServiceAccountDescription}
                    disabled={!canPersistSettings}
                  />
                  {(draftSettings.emailServiceAccount || "").trim() && (
                    <SharePointButton
                      variant="secondary"
                      disabled={!canPersistSettings}
                      onClick={async () => {
                        try {
                          if (!context?.msGraphClientFactory) {
                            notificationService?.showError(
                              strings.SettingsGraphUnavailableError,
                              strings.Error,
                            );
                            return;
                          }
                          const graphClient =
                            await context.msGraphClientFactory.getClient("3");
                          const emailService = new EmailNotificationService({
                            serviceAccountEmail:
                              draftSettings.emailServiceAccount || "",
                            graphClient,
                          });
                          await emailService.sendTestEmail(
                            context?.pageContext?.user?.email || "",
                          );
                          notificationService?.showSuccess(
                            strings.SettingsTestEmailSuccess,
                            strings.Success,
                          );
                        } catch (error) {
                          logger.error(
                            "SettingsTab",
                            "Test email failed",
                            error,
                          );
                          notificationService?.showError(
                            strings.SettingsTestEmailFailed,
                            strings.Error,
                          );
                        }
                      }}
                    >
                      {strings.SettingsSendTestEmailButton}
                    </SharePointButton>
                  )}
                </div>
              )}

              <SharePointToggle
                label={strings.EnableCopilotLabel}
                checked={draftSettings.copilotEnabled || false}
                onChange={(checked) =>
                  handleSettingsChange({ copilotEnabled: checked })
                }
                description={strings.EnableCopilotDescription}
                disabled={!canPersistSettings}
              />
            </div>
          </SharePointSection>
        </div>

        <SharePointSection title={strings.SettingsSectionExperienceTitle}>
          <div className={styles.settingsGrid}>
            <SharePointToggle
              label={strings.SettingsCarouselEnabledLabel}
              checked={carouselEnabled}
              onChange={handleCarouselEnabledChange}
              description={strings.SettingsCarouselEnabledDescription}
              disabled={!canPersistSettings}
            />

            <SharePointInput
              label={strings.SettingsCarouselIntervalLabel}
              value={carouselInterval}
              onChange={handleCarouselIntervalChange}
              placeholder="5"
              type="text"
              description={strings.SettingsCarouselIntervalDescription}
              disabled={!carouselEnabled || !canPersistSettings}
            />

            <SharePointSelect
              label="Alert Sort Order"
              value={alertsState.sortMode}
              onChange={(value) => updateSortMode(value as AlertSortMode)}
              options={[
                { value: "priority", label: "Priority (Critical → Low)" },
                { value: "date", label: "Date (Newest First)" },
                { value: "alphabetical", label: "Alphabetical (A → Z)" },
                { value: "manual", label: "Manual (Custom Order)" },
              ]}
              description="Choose how alerts are ordered in the banner"
            />
          </div>
        </SharePointSection>

        <SharePointSection title="Priority Border Colors">
          <p className={styles.fieldDescription}>
            Customize the left border color for each priority level.
          </p>
          <div className={styles.priorityColorsGrid}>
            {(
              [
                {
                  priority: AlertPriority.Critical,
                  label: "Critical",
                  defaultColor: "#d13438",
                },
                {
                  priority: AlertPriority.High,
                  label: "High",
                  defaultColor: "#f7630c",
                },
                {
                  priority: AlertPriority.Medium,
                  label: "Medium",
                  defaultColor: "#0078d4",
                },
                {
                  priority: AlertPriority.Low,
                  label: "Low",
                  defaultColor: "#107c10",
                },
              ] as const
            ).map(({ priority, label, defaultColor }) => (
              <div key={priority} className={styles.priorityColorItem}>
                <ColorPicker
                  label={label}
                  value={
                    priorityBorderColors[priority]?.borderColor || defaultColor
                  }
                  onChange={(color) =>
                    setPriorityBorderColors((prev) => ({
                      ...prev,
                      [priority]: { borderColor: color },
                    }))
                  }
                  description=""
                />
              </div>
            ))}
          </div>
        </SharePointSection>

        <div id="settings-section-setup">
          {(alertsListExists === false || alertTypesListExists === false) && (
            <SharePointSection title={strings.SettingsSectionSetupTitle}>
              <div className={styles.settingsGrid}>
                <div className={styles.fullWidthColumn}>
                  {isCheckingLists ? (
                    <div className={styles.spinnerContainer}>
                      <div className={styles.spinner}></div>
                      {strings.SettingsCheckingLists}
                    </div>
                  ) : (
                    <>
                      <p className={styles.infoText}>
                        {strings.SettingsMissingListsDescription}
                      </p>
                      <div className={styles.infoText}>
                        <strong>{strings.SettingsCurrentSiteLabel}</strong>{" "}
                        {window.location.href.split("/")[2]}
                      </div>
                      <ul className={styles.infoText}>
                        {alertsListExists === false && (
                          <li>{strings.SettingsMissingAlertsListItem}</li>
                        )}
                        {alertTypesListExists === false && (
                          <li>{strings.SettingsMissingTypesListItem}</li>
                        )}
                      </ul>

                      <Card>
                        <CardHeader
                          header={
                            <div className={cardStyles.cardHeader}>
                              <Globe24Regular />
                              <Text weight="semibold">
                                {strings.SettingsInitialLanguagesTitle}
                              </Text>
                            </div>
                          }
                          description={
                            <Text size={200}>
                              {strings.SettingsInitialLanguagesDescription}
                            </Text>
                          }
                        />

                        <CardPreview>
                          <div className={cardStyles.cardContent}>
                            <div className={cardStyles.languageGrid}>
                              {LanguageAwarenessService.getSupportedLanguages().map(
                                (language) => (
                                  <div
                                    key={language.code}
                                    className={cardStyles.languageItem}
                                  >
                                    <Checkbox
                                      checked={preCreationLanguages.includes(
                                        language.code,
                                      )}
                                      disabled={language.code === "en-us"}
                                      onChange={(_, data) => {
                                        if (data.checked === true) {
                                          setPreCreationLanguages((prev) => [
                                            ...prev,
                                            language.code,
                                          ]);
                                        } else if (language.code !== "en-us") {
                                          setPreCreationLanguages((prev) =>
                                            prev.filter(
                                              (code) => code !== language.code,
                                            ),
                                          );
                                        }
                                      }}
                                    />
                                    <div className={cardStyles.languageInfo}>
                                      <div className={cardStyles.languageName}>
                                        {language.flag} {language.nativeName}
                                      </div>
                                      <div className={cardStyles.languageCode}>
                                        {language.name} (
                                        {language.code.toUpperCase()})
                                      </div>
                                    </div>
                                  </div>
                                ),
                              )}
                            </div>
                            <Text size={100} className={cardStyles.hintText}>
                              {strings.SettingsInitialLanguagesHint}
                            </Text>
                          </div>
                        </CardPreview>
                      </Card>

                      <div className={styles.actionButtonsRow}>
                        <SharePointButton
                          variant="primary"
                          icon={<Add24Regular />}
                          onClick={handleCreateLists}
                          disabled={isCreatingLists || !canPersistSettings}
                        >
                          {isCreatingLists
                            ? strings.SettingsCreatingLists
                            : strings.SettingsCreateMissingLists}
                        </SharePointButton>

                        <div className={styles.helpText}>
                          {strings.SettingsCreateMissingListsHelp}
                        </div>
                      </div>

                      {isCreatingLists && creationSteps.length > 0 && (
                        <div className={styles.creatingProgress}>
                          <ProgressIndicator
                            steps={creationSteps}
                            title={strings.SettingsCreatingLists}
                            showStepDescriptions={true}
                            variant="vertical"
                          />
                        </div>
                      )}
                    </>
                  )}
                </div>
              </div>
            </SharePointSection>
          )}

          {alertsListExists === true && alertTypesListExists === true && (
            <SharePointSection title={strings.SettingsSectionSetupTitle}>
              <div className={styles.successContainer}>
                <div className={styles.successHeader}>
                  <span className={styles.successIcon}>✅</span>
                  <strong>{strings.SettingsSetupCompleteTitle}</strong>
                </div>
                <p className={styles.successDescription}>
                  {strings.SettingsSetupCompleteDescription}
                </p>

                <div className={styles.additionalOptions}>
                  <h4>{strings.SettingsListMaintenanceTitle}</h4>
                  <div className={styles.actionButtonsRow}>
                    <SharePointButton
                      variant="secondary"
                      icon={<Wrench24Regular />}
                      onClick={handleOpenRepairDialog}
                      disabled={!canPersistSettings}
                    >
                      {strings.RepairDialogTitle}
                    </SharePointButton>
                    <div className={styles.helpText}>
                      {strings.SettingsListMaintenanceDescription}
                    </div>
                  </div>
                </div>
              </div>
            </SharePointSection>
          )}
        </div>

        {alertsListExists === true && alertTypesListExists === true && (
          <div id="settings-section-localization">
            <SharePointSection title={strings.SettingsSectionLocalizationTitle}>
              <div className={styles.additionalOptions}>
                <h3 className={styles.languageManagementTitle}>
                  <LocalLanguage24Regular
                    className={styles.languageDialogIcon}
                  />
                  {strings.ManageLanguagesForList}
                </h3>
                <p className={styles.languageManagementDescription}>
                  {strings.LanguageManagerDescription}
                </p>
                <LanguageFieldManager
                  alertService={alertService}
                  onLanguageChange={onLanguageChange}
                />

                <Card>
                  <CardHeader
                    header={
                      <div className={cardStyles.cardHeader}>
                        <Globe24Regular />
                        <Text weight="semibold">
                          {strings.LanguagePolicyTitle}
                        </Text>
                      </div>
                    }
                    description={
                      <Text size={200}>
                        {strings.LanguagePolicyDescription}
                      </Text>
                    }
                  />
                  <CardPreview>
                    <div className={cardStyles.cardContent}>
                      <div className={styles.settingsGrid}>
                        <SharePointSelect
                          label={strings.LanguagePolicyFallbackLanguageLabel}
                          value={languagePolicy.fallbackLanguage}
                          onChange={(value) =>
                            setLanguagePolicy((prev) => ({
                              ...prev,
                              fallbackLanguage: value as any,
                            }))
                          }
                          options={fallbackOptions}
                          description={
                            strings.LanguagePolicyFallbackLanguageDescription
                          }
                          disabled={!canPersistSettings}
                        />
                        <SharePointSelect
                          label={strings.LanguagePolicyCompletenessRuleLabel}
                          value={languagePolicy.completenessRule}
                          onChange={(value) =>
                            setLanguagePolicy((prev) => ({
                              ...prev,
                              completenessRule: value as any,
                            }))
                          }
                          options={completenessOptions}
                          description={
                            strings.LanguagePolicyCompletenessRuleDescription
                          }
                          disabled={!canPersistSettings}
                        />
                        <SharePointToggle
                          label={
                            strings.LanguagePolicyRequireLinkDescriptionLabel
                          }
                          checked={languagePolicy.requireLinkDescriptionWhenUrl}
                          onChange={(checked) =>
                            setLanguagePolicy((prev) => ({
                              ...prev,
                              requireLinkDescriptionWhenUrl: checked,
                            }))
                          }
                          description={
                            strings.LanguagePolicyRequireLinkDescriptionDescription
                          }
                          disabled={!canPersistSettings}
                        />
                      </div>

                      {showAdvancedLocalization && (
                        <>
                          <div className={styles.settingsGrid}>
                            <SharePointToggle
                              label={strings.LanguagePolicyInheritanceEnable}
                              checked={languagePolicy.inheritance.enabled}
                              onChange={(checked) =>
                                setLanguagePolicy((prev) => ({
                                  ...prev,
                                  inheritance: {
                                    ...prev.inheritance,
                                    enabled: checked,
                                  },
                                }))
                              }
                              description={
                                strings.LanguagePolicyInheritanceDescription
                              }
                              disabled={!canPersistSettings}
                            />

                            {languagePolicy.inheritance.enabled && (
                              <>
                                <Checkbox
                                  checked={
                                    languagePolicy.inheritance.fields.title
                                  }
                                  label={
                                    strings.LanguagePolicyInheritanceTitleField
                                  }
                                  onChange={(_, data) =>
                                    setLanguagePolicy((prev) => ({
                                      ...prev,
                                      inheritance: {
                                        ...prev.inheritance,
                                        fields: {
                                          ...prev.inheritance.fields,
                                          title: !!data.checked,
                                        },
                                      },
                                    }))
                                  }
                                />
                                <Checkbox
                                  checked={
                                    languagePolicy.inheritance.fields
                                      .description
                                  }
                                  label={
                                    strings.LanguagePolicyInheritanceDescriptionField
                                  }
                                  onChange={(_, data) =>
                                    setLanguagePolicy((prev) => ({
                                      ...prev,
                                      inheritance: {
                                        ...prev.inheritance,
                                        fields: {
                                          ...prev.inheritance.fields,
                                          description: !!data.checked,
                                        },
                                      },
                                    }))
                                  }
                                />
                                <Checkbox
                                  checked={
                                    languagePolicy.inheritance.fields
                                      .linkDescription
                                  }
                                  label={
                                    strings.LanguagePolicyInheritanceLinkDescriptionField
                                  }
                                  onChange={(_, data) =>
                                    setLanguagePolicy((prev) => ({
                                      ...prev,
                                      inheritance: {
                                        ...prev.inheritance,
                                        fields: {
                                          ...prev.inheritance.fields,
                                          linkDescription: !!data.checked,
                                        },
                                      },
                                    }))
                                  }
                                />
                              </>
                            )}
                          </div>

                          <div className={styles.settingsGrid}>
                            <SharePointToggle
                              label={strings.LanguagePolicyWorkflowEnable}
                              checked={languagePolicy.workflow.enabled}
                              onChange={(checked) =>
                                setLanguagePolicy((prev) => ({
                                  ...prev,
                                  workflow: {
                                    ...prev.workflow,
                                    enabled: checked,
                                  },
                                }))
                              }
                              description={
                                strings.LanguagePolicyWorkflowDescription
                              }
                              disabled={!canPersistSettings}
                            />

                            {languagePolicy.workflow.enabled && (
                              <>
                                <SharePointSelect
                                  label={
                                    strings.LanguagePolicyWorkflowDefaultStatusLabel
                                  }
                                  value={languagePolicy.workflow.defaultStatus}
                                  onChange={(value) =>
                                    setLanguagePolicy((prev) => ({
                                      ...prev,
                                      workflow: {
                                        ...prev.workflow,
                                        defaultStatus:
                                          value as TranslationStatus,
                                      },
                                    }))
                                  }
                                  options={workflowStatusOptions}
                                  disabled={!canPersistSettings}
                                />
                                <SharePointToggle
                                  label={
                                    strings.LanguagePolicyWorkflowRequireApproved
                                  }
                                  checked={
                                    languagePolicy.workflow
                                      .requireApprovedForDisplay
                                  }
                                  onChange={(checked) =>
                                    setLanguagePolicy((prev) => ({
                                      ...prev,
                                      workflow: {
                                        ...prev.workflow,
                                        requireApprovedForDisplay: checked,
                                      },
                                    }))
                                  }
                                  description={
                                    strings.LanguagePolicyWorkflowRequireApprovedDescription
                                  }
                                  disabled={!canPersistSettings}
                                />
                              </>
                            )}
                          </div>
                        </>
                      )}

                      <div className={styles.sectionActionToolbar}>
                        <div className={styles.sectionActionPrimary}>
                          <SharePointButton
                            variant="secondary"
                            onClick={() =>
                              setShowAdvancedLocalization((prev) => !prev)
                            }
                            className={styles.actionToggleButton}
                          >
                            {showAdvancedLocalization
                              ? strings.SettingsHideAdvanced
                              : strings.SettingsShowAdvanced}
                          </SharePointButton>
                        </div>
                        <div className={styles.sectionActionSecondary}>
                          <SharePointButton
                            variant="primary"
                            onClick={handleSaveLanguagePolicy}
                            disabled={isSavingPolicy || !canPersistSettings}
                            className={styles.actionPrimaryButton}
                          >
                            {isSavingPolicy
                              ? strings.LanguagePolicySaving
                              : strings.LanguagePolicySaveButton}
                          </SharePointButton>
                        </div>
                      </div>
                    </div>
                  </CardPreview>
                </Card>
              </div>
            </SharePointSection>
          </div>
        )}

        <div id="settings-section-maintenance">
          <SharePointSection title={strings.SettingsSectionMaintenanceTitle}>
            <div className={styles.settingsGrid}>
              <div className={styles.fullWidthColumn}>
                <p className={styles.storageManagement}>
                  {strings.SettingsStorageDescription}
                </p>
                <div className={styles.storageButtons}>
                  <SharePointButton
                    variant="secondary"
                    onClick={() => {
                      storageService.current.clearAllAlertData();
                      notificationService?.showSuccess(
                        strings.SettingsCacheCleared,
                        strings.Success,
                      );
                    }}
                  >
                    {strings.SettingsClearCache}
                  </SharePointButton>
                  <SharePointButton
                    variant="secondary"
                    onClick={() => {
                      syncCarouselToContext(false, 5);
                      setCarouselEnabled(false);
                      setCarouselInterval("5");
                      notificationService?.showSuccess(
                        strings.SettingsCarouselResetSuccess,
                        strings.Success,
                      );
                    }}
                  >
                    {strings.SettingsResetCarousel}
                  </SharePointButton>
                </div>

                <div className={styles.sectionActionToolbar}>
                  <div className={styles.sectionActionPrimary}>
                    <SharePointButton
                      variant="secondary"
                      onClick={() =>
                        setShowAdvancedMaintenance((prev) => !prev)
                      }
                      className={styles.actionToggleButton}
                    >
                      {showAdvancedMaintenance
                        ? strings.SettingsHideAdvanced
                        : strings.SettingsShowAdvanced}
                    </SharePointButton>
                  </div>
                  {showAdvancedMaintenance && (
                    <div
                      className={`${styles.sectionActionSecondary} ${styles.storageButtons}`}
                    >
                      <SharePointButton
                        variant="secondary"
                        onClick={handleExportSettings}
                      >
                        {strings.SettingsExportButton}
                      </SharePointButton>
                      <SharePointButton
                        variant="secondary"
                        onClick={handleImportSettingsClick}
                      >
                        {strings.SettingsImportButton}
                      </SharePointButton>
                      <SharePointButton
                        variant="secondary"
                        onClick={handleResetToDefaults}
                        disabled={!canPersistSettings}
                      >
                        {strings.SettingsResetToDefaults}
                      </SharePointButton>
                    </div>
                  )}
                </div>
                <input
                  ref={importInputRef}
                  type="file"
                  accept=".json,application/json"
                  onChange={handleImportSettingsFile}
                  className={styles.hiddenFileInput}
                />
              </div>
            </div>
          </SharePointSection>
        </div>

        <SharePointSection title={strings.SettingsAboutTitle}>
          <div className={styles.aboutSection}>
            <div className={styles.aboutCard}>
              <p className={styles.aboutTitle}>
                {strings.SettingsAboutProductName}
              </p>
              <p className={styles.aboutVersion}>
                {CoreText.format(strings.SettingsAboutVersion, "5.0.1")}
              </p>
              <p className={styles.aboutDescription}>
                {strings.SettingsAboutDescription}
              </p>
              <p className={styles.aboutAuthor}>
                {strings.SettingsAboutAuthor}
              </p>
              <a
                href={strings.SettingsAboutGitHubUrl}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.aboutLink}
              >
                {strings.SettingsAboutGitHubLink}
                <ArrowRight24Regular />
              </a>
            </div>
          </div>
        </SharePointSection>
      </div>

      <RepairDialog
        isOpen={isRepairDialogOpen}
        onDismiss={handleCloseRepairDialog}
        onRepairComplete={handleRepairComplete}
        alertService={alertService}
      />
    </div>
  );
};

export default SettingsTab;
