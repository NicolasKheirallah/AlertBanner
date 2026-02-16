import * as React from "react";
import { Settings24Regular, Add24Regular } from "@fluentui/react-icons";
import SharePointDialog from "../UI/SharePointDialog";
import { SharePointButton } from "../UI/SharePointControls";
import CreateAlertTab, { INewAlert, IFormErrors } from "./Tabs/CreateAlertTab";
import ManageAlertsTab, { IEditingAlert } from "./Tabs/ManageAlertsTab";
import AlertTypesTab from "./Tabs/AlertTypesTab";
import SettingsTab, { ISettingsData } from "./Tabs/SettingsTab";
import {
  AlertPriority,
  NotificationType,
  IAlertType,
  ContentType,
  TargetLanguage,
} from "../Alerts/IAlerts";
import {
  SiteContextDetector,
  ISiteValidationResult,
} from "../Utils/SiteContextDetector";
import { SharePointAlertService } from "../Services/SharePointAlertService";
import { IAlertItem } from "../Alerts/IAlerts";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "./AlertSettings.module.scss";
import { logger } from "../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { useFluentDialogs } from "../Hooks/useFluentDialogs";

export interface IAlertSettingsTabsProps {
  isInEditMode: boolean;
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  emailServiceAccount?: string;
  copilotEnabled?: boolean;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  onSettingsChange: (
    settings: ISettingsData & { enableTargetSite: boolean },
  ) => void;
}

const AlertSettingsTabs: React.FC<IAlertSettingsTabsProps> = ({
  isInEditMode,
  alertTypesJson,
  userTargetingEnabled,
  notificationsEnabled,
  enableTargetSite,
  emailServiceAccount,
  copilotEnabled,
  graphClient,
  context,
  onSettingsChange,
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<
    "create" | "manage" | "types" | "settings"
  >("create");

  // Shared services - using useRef to prevent recreation
  const siteDetector = React.useRef<SiteContextDetector>(
    new SiteContextDetector(graphClient, context),
  );
  const alertService = React.useRef<SharePointAlertService>(
    new SharePointAlertService(graphClient, context),
  );
  const [languageUpdateTrigger, setLanguageUpdateTrigger] = React.useState(0);

  // Site context (removed unused variable)

  // Settings state
  const [settings, setSettings] = React.useState<
    ISettingsData & { enableTargetSite: boolean }
  >({
    alertTypesJson,
    userTargetingEnabled,
    notificationsEnabled,
    enableTargetSite,
    emailServiceAccount,
    copilotEnabled: !!copilotEnabled,
  });

  // Alert types state
  const [alertTypes, setAlertTypes] = React.useState<IAlertType[]>([]);

  // Create alert state
  const [newAlert, setNewAlert] = React.useState<INewAlert>({
    title: "",
    description: "",
    AlertType: "",
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
  });
  const [errors, setErrors] = React.useState<IFormErrors>({});
  const [creationProgress, setCreationProgress] = React.useState<
    ISiteValidationResult[]
  >([]);
  const [isCreatingAlert, setIsCreatingAlert] = React.useState(false);
  const [showPreview, setShowPreview] = React.useState(true);
  const [showTemplates, setShowTemplates] = React.useState(true);
  const [hasUnsavedCreateChanges, setHasUnsavedCreateChanges] =
    React.useState(false);
  const [hasUnsavedManageChanges, setHasUnsavedManageChanges] =
    React.useState(false);
  const [hasUnsavedSettingsChanges, setHasUnsavedSettingsChanges] =
    React.useState(false);

  // Manage alerts state
  const [existingAlerts, setExistingAlerts] = React.useState<IAlertItem[]>([]);
  const [isLoadingAlerts, setIsLoadingAlerts] = React.useState(false);
  const [selectedAlerts, setSelectedAlerts] = React.useState<string[]>([]);
  const [editingAlert, setEditingAlert] = React.useState<IEditingAlert | null>(
    null,
  );
  const [isEditingAlert, setIsEditingAlert] = React.useState(false);

  // Alert types state
  const [newAlertType, setNewAlertType] = React.useState<IAlertType>({
    name: "",
    iconName: "Info",
    backgroundColor: "#0078d4",
    textColor: "#ffffff",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: "",
    },
  });
  const [isCreatingType, setIsCreatingType] = React.useState(false);

  // SharePoint list state
  const [alertsListExists, setAlertsListExists] = React.useState<
    boolean | null
  >(null);
  const [alertTypesListExists, setAlertTypesListExists] = React.useState<
    boolean | null
  >(null);
  const [isCheckingLists, setIsCheckingLists] = React.useState(false);
  const [isCreatingLists, setIsCreatingLists] = React.useState(false);
  const alertTypesLoadInFlightRef = React.useRef<Promise<void> | null>(null);
  const { confirm, dialogs } = useFluentDialogs();

  const buildInitialNewAlert = React.useCallback((): INewAlert => {
    return {
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
    };
  }, [alertTypes]);

  const resetCreateTabState = React.useCallback(() => {
    setShowTemplates(true);
    setShowPreview(true);
    setHasUnsavedCreateChanges(false);
    setErrors({});
    setCreationProgress([]);
    setIsCreatingAlert(false);
    setNewAlert(buildInitialNewAlert());
  }, [buildInitialNewAlert]);

  const openNewCreateAlert = React.useCallback(() => {
    setEditingAlert(null);
    setIsEditingAlert(false);
    setHasUnsavedManageChanges(false);
    setActiveTab("create");
    resetCreateTabState();
  }, [resetCreateTabState]);

  const confirmDiscardCreateChanges = React.useCallback(async (): Promise<boolean> => {
    if (!hasUnsavedCreateChanges || activeTab !== "create") {
      return true;
    }

    return confirm({
      title: strings.CreateAlertUnsavedChangesTitle,
      message: strings.CreateAlertUnsavedChangesMessage,
      confirmText: strings.CreateAlertDiscardChangesButton,
      cancelText: strings.CreateAlertKeepEditingButton,
    });
  }, [activeTab, confirm, hasUnsavedCreateChanges]);

  const confirmDiscardManageChanges = React.useCallback(
    async (): Promise<boolean> => {
      if (!hasUnsavedManageChanges || activeTab !== "manage") {
        return true;
      }

      return confirm({
        title: strings.ManageAlertsUnsavedChangesTitle,
        message: strings.ManageAlertsUnsavedChangesMessage,
        confirmText: strings.ManageAlertsDiscardChangesButton,
        cancelText: strings.ManageAlertsKeepEditingButton,
      });
    },
    [activeTab, confirm, hasUnsavedManageChanges],
  );

  const confirmDiscardSettingsChanges = React.useCallback(
    async (): Promise<boolean> => {
      if (!hasUnsavedSettingsChanges || activeTab !== "settings") {
        return true;
      }

      return confirm({
        title: strings.SettingsUnsavedChangesTitle,
        message: strings.SettingsUnsavedChangesMessage,
        confirmText: strings.SettingsDiscardChanges,
        cancelText: strings.SettingsKeepEditing,
      });
    },
    [activeTab, confirm, hasUnsavedSettingsChanges],
  );

  const switchTab = React.useCallback(
    async (nextTab: "create" | "manage" | "types" | "settings") => {
      if (nextTab === activeTab) {
        return;
      }

      if (nextTab !== "create") {
        const canLeave = await confirmDiscardCreateChanges();
        if (!canLeave) {
          return;
        }
      }
      if (nextTab !== "manage") {
        const canLeave = await confirmDiscardManageChanges();
        if (!canLeave) {
          return;
        }
      }
      if (nextTab !== "settings") {
        const canLeave = await confirmDiscardSettingsChanges();
        if (!canLeave) {
          return;
        }
      }

      if (nextTab === "create") {
        openNewCreateAlert();
        return;
      }

      setActiveTab(nextTab);
    },
    [
      activeTab,
      confirmDiscardCreateChanges,
      confirmDiscardManageChanges,
      confirmDiscardSettingsChanges,
      openNewCreateAlert,
    ],
  );

  const handleCloseDialog = React.useCallback(async () => {
    const canCloseCreate = await confirmDiscardCreateChanges();
    if (!canCloseCreate) {
      return;
    }
    const canCloseManage = await confirmDiscardManageChanges();
    if (!canCloseManage) {
      return;
    }
    const canCloseSettings = await confirmDiscardSettingsChanges();
    if (!canCloseSettings) {
      return;
    }

    if (editingAlert) {
      setEditingAlert(null);
      setIsEditingAlert(false);
    }
    openNewCreateAlert();
    setIsOpen(false);
  }, [
    confirmDiscardCreateChanges,
    confirmDiscardManageChanges,
    confirmDiscardSettingsChanges,
    editingAlert,
    openNewCreateAlert,
  ]);

  const loadAlertTypes = React.useCallback(async () => {
    if (alertTypesLoadInFlightRef.current) {
      await alertTypesLoadInFlightRef.current;
      return;
    }

    const task = (async () => {
      try {
        const types = await alertService.current.getAlertTypes();
        if (!types || types.length === 0) {
          setAlertTypes([]);
          return;
        }

        setAlertTypes(types);

        // Set first alert type as default if none is selected
        setNewAlert((prev) => {
          if (
            !prev.AlertType ||
            !types.find((t) => t.name === prev.AlertType)
          ) {
            return { ...prev, AlertType: types[0].name };
          }
          return prev;
        });
      } catch (error) {
        logger.error(
          "AlertSettingsTabs",
          "Error loading alert types from SharePoint",
          error,
        );
        setAlertTypes([]);
      }
    })();

    alertTypesLoadInFlightRef.current = task;
    await task;
    alertTypesLoadInFlightRef.current = null;
  }, []);

  React.useEffect(() => {
    if (!isOpen) {
      return;
    }

    loadAlertTypes();
  }, [isOpen, loadAlertTypes]);

  // Initialize site context
  React.useEffect(() => {
    if (!isOpen) {
      return;
    }

    siteDetector.current
      .getCurrentSiteContext()
      .then((siteContext) => {
        // Set current site as default target if no sites selected
        if (newAlert.targetSites.length === 0) {
          setNewAlert((prev) => ({
            ...prev,
            targetSites: [siteContext.siteId],
          }));
        }
      })
      .catch((error) => {
        logger.error("AlertSettingsTabs", "Failed to get site context", error);
      });
  }, [isOpen, newAlert.targetSites.length]);

  // Update settings when props change
  React.useEffect(() => {
    setSettings({
      alertTypesJson,
      userTargetingEnabled,
      notificationsEnabled,
      enableTargetSite,
      emailServiceAccount,
      copilotEnabled: !!copilotEnabled,
    });
  }, [
    alertTypesJson,
    userTargetingEnabled,
    notificationsEnabled,
    enableTargetSite,
    emailServiceAccount,
    copilotEnabled,
  ]);

  const handleSettingsChange = React.useCallback(
    (newSettings: ISettingsData) => {
      setSettings(newSettings);
      onSettingsChange(newSettings);
    },
    [onSettingsChange],
  );

  const handleLanguageChange = React.useCallback((languages: string[]) => {
    logger.debug("AlertSettingsTabs", "Languages changed, triggering refresh", {
      languages,
    });
    setLanguageUpdateTrigger((prev) => prev + 1);
  }, []);

  const canManageSettings =
    isInEditMode ||
    !!context?.pageContext?.legacyPageContext?.isSiteAdmin;

  if (!canManageSettings) {
    return null;
  }

  return (
    <>
      <div className={styles.settingsButton}>
        <SharePointButton
          variant="secondary"
          icon={<Settings24Regular />}
          onClick={() => {
            openNewCreateAlert();
            setIsOpen(true);
          }}
          aria-label={strings.AlertSettingsTitle}
          title={strings.AlertSettingsTitle}
        />
      </div>

      <SharePointDialog
        isOpen={isOpen}
        onClose={() => {
          void handleCloseDialog();
        }}
        title={
          editingAlert
            ? `${strings.EditAlert}: ${editingAlert.title}`
            : strings.AlertSettingsTitle
        }
        width={1200}
      >
        <div className={styles.settingsContainer}>
          {/* Tab Navigation — hidden when editing an alert */}
          {!editingAlert && (
            <div
              className={styles.tabs}
              role="tablist"
              aria-label={strings.AlertSettingsTitle}
            >
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  void switchTab("create");
                }}
                className={`${styles.tab} ${activeTab === "create" ? styles.activeTab : ""}`}
                icon={<Add24Regular />}
                role="tab"
                aria-selected={activeTab === "create"}
                aria-controls="tabpanel-create"
                id="tab-create"
              >
                {strings.CreateAlert}
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  void switchTab("manage");
                }}
                className={`${styles.tab} ${activeTab === "manage" ? styles.activeTab : ""}`}
                role="tab"
                aria-selected={activeTab === "manage"}
                aria-controls="tabpanel-manage"
                id="tab-manage"
              >
                {strings.ManageAlerts}
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  void switchTab("types");
                }}
                className={`${styles.tab} ${activeTab === "types" ? styles.activeTab : ""}`}
                role="tab"
                aria-selected={activeTab === "types"}
                aria-controls="tabpanel-types"
                id="tab-types"
              >
                {strings.AlertTypesTabTitle}
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  void switchTab("settings");
                }}
                className={`${styles.tab} ${activeTab === "settings" ? styles.activeTab : ""}`}
                icon={<Settings24Regular />}
                role="tab"
                aria-selected={activeTab === "settings"}
                aria-controls="tabpanel-settings"
                id="tab-settings"
              >
                {strings.SettingsTabTitle}
              </SharePointButton>
            </div>
          )}

          {/* Tab Content */}
          <div className={styles.tabContent}>
            {activeTab === "create" && (
              <div
                role="tabpanel"
                id="tabpanel-create"
                aria-labelledby="tab-create"
              >
                <CreateAlertTab
                  newAlert={newAlert}
                  setNewAlert={setNewAlert}
                  errors={errors}
                  setErrors={setErrors}
                  alertTypes={alertTypes}
                  userTargetingEnabled={userTargetingEnabled}
                  notificationsEnabled={notificationsEnabled}
                  enableTargetSite={settings.enableTargetSite}
                  siteDetector={siteDetector.current}
                  alertService={alertService.current}
                  graphClient={graphClient}
                  context={context}
                  creationProgress={creationProgress}
                  setCreationProgress={setCreationProgress}
                  isCreatingAlert={isCreatingAlert}
                  setIsCreatingAlert={setIsCreatingAlert}
                  showPreview={showPreview}
                  setShowPreview={setShowPreview}
                  showTemplates={showTemplates}
                  setShowTemplates={setShowTemplates}
                  languageUpdateTrigger={languageUpdateTrigger}
                  copilotEnabled={settings.copilotEnabled}
                  onDirtyStateChange={setHasUnsavedCreateChanges}
                />
              </div>
            )}

            {activeTab === "manage" && (
              <div
                role="tabpanel"
                id="tabpanel-manage"
                aria-labelledby="tab-manage"
              >
                <ManageAlertsTab
                  existingAlerts={existingAlerts}
                  setExistingAlerts={setExistingAlerts}
                  isLoadingAlerts={isLoadingAlerts}
                  setIsLoadingAlerts={setIsLoadingAlerts}
                  selectedAlerts={selectedAlerts}
                  setSelectedAlerts={setSelectedAlerts}
                  editingAlert={editingAlert}
                  setEditingAlert={setEditingAlert}
                  isEditingAlert={isEditingAlert}
                  setIsEditingAlert={setIsEditingAlert}
                  alertTypes={alertTypes}
                  siteDetector={siteDetector.current}
                  alertService={alertService.current}
                  graphClient={graphClient}
                  context={context}
                  userTargetingEnabled={settings.userTargetingEnabled}
                  notificationsEnabled={settings.notificationsEnabled}
                  enableTargetSite={settings.enableTargetSite}
                  copilotEnabled={settings.copilotEnabled}
                  onDirtyStateChange={setHasUnsavedManageChanges}
                  setActiveTab={(nextTab) => {
                    if (typeof nextTab === "function") {
                      const resolved = nextTab(activeTab);
                      void switchTab(resolved);
                      return;
                    }

                    void switchTab(nextTab);
                  }}
                />
              </div>
            )}

            {activeTab === "types" && (
              <div
                role="tabpanel"
                id="tabpanel-types"
                aria-labelledby="tab-types"
              >
                <AlertTypesTab
                  alertTypes={alertTypes}
                  setAlertTypes={setAlertTypes}
                  newAlertType={newAlertType}
                  setNewAlertType={setNewAlertType}
                  isCreatingType={isCreatingType}
                  setIsCreatingType={setIsCreatingType}
                  alertService={alertService.current}
                  context={context}
                />
              </div>
            )}

            {activeTab === "settings" && (
              <div
                role="tabpanel"
                id="tabpanel-settings"
                aria-labelledby="tab-settings"
              >
                <SettingsTab
                  settings={settings}
                  setSettings={setSettings}
                  alertsListExists={alertsListExists}
                  setAlertsListExists={setAlertsListExists}
                  alertTypesListExists={alertTypesListExists}
                  setAlertTypesListExists={setAlertTypesListExists}
                  isCheckingLists={isCheckingLists}
                  setIsCheckingLists={setIsCheckingLists}
                  isCreatingLists={isCreatingLists}
                  setIsCreatingLists={setIsCreatingLists}
                  alertService={alertService.current}
                  onSettingsChange={handleSettingsChange}
                  onLanguageChange={handleLanguageChange}
                  onDirtyStateChange={setHasUnsavedSettingsChanges}
                  canEdit={canManageSettings}
                  context={context}
                />
              </div>
            )}
          </div>
        </div>
      </SharePointDialog>
      {dialogs}
    </>
  );
};

export default AlertSettingsTabs;
