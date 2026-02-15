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
import { useAsyncOperation } from "../Hooks/useAsyncOperation";
import * as strings from "AlertBannerApplicationCustomizerStrings";

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

  // Initialize alert types from SharePoint using useAsyncOperation
  const { execute: loadAlertTypes } = useAsyncOperation(
    async () => {
      const types = await alertService.current.getAlertTypes();
      return types;
    },
    {
      onSuccess: (types) => {
        if (types && types.length > 0) {
          setAlertTypes(types);

          // Set first alert type as default if none is selected
          setNewAlert((prev) => {
            // Only set default if AlertType is empty or invalid
            if (
              !prev.AlertType ||
              !types.find((t) => t.name === prev.AlertType)
            ) {
              return { ...prev, AlertType: types[0].name };
            }
            return prev;
          });
        }
      },
      onError: () => {
        logger.error(
          "AlertSettingsTabs",
          "Error loading alert types from SharePoint",
        );
        setAlertTypes([]);
      },
      logErrors: true,
    },
  );

  React.useEffect(
    () => {
      // Edit mode guard disabled — always load alert types
      // if (isInEditMode) {
      loadAlertTypes();
      // }
    },
    [
      /* isInEditMode */
    ],
  );

  // Initialize site context
  React.useEffect(() => {
    // Edit mode guard disabled — always initialize site context
    // if (isInEditMode) {
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
    // }
  }, [/* isInEditMode, */ newAlert.targetSites.length]);

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

  // Edit mode guard disabled — always render settings button
  // if (!isInEditMode) {
  //   return null;
  // }

  return (
    <>
      <div className={styles.settingsButton}>
        <SharePointButton
          variant="secondary"
          icon={<Settings24Regular />}
          onClick={() => setIsOpen(true)}
        />
      </div>

      <SharePointDialog
        isOpen={isOpen}
        onClose={() => {
          if (editingAlert) {
            setEditingAlert(null);
            setIsEditingAlert(false);
          }
          setIsOpen(false);
        }}
        title={
          editingAlert
            ? `Edit ${
                editingAlert.contentType === ContentType.Template
                  ? "Template"
                  : editingAlert.contentType === ContentType.Draft
                    ? "Draft"
                    : "Alert"
              }: ${editingAlert.title}`
            : strings.AlertSettingsTitle
        }
        width={1200}
        height={800}
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
                onClick={() => setActiveTab("create")}
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
                onClick={() => setActiveTab("manage")}
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
                onClick={() => setActiveTab("types")}
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
                onClick={() => setActiveTab("settings")}
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
                  copilotEnabled={settings.copilotEnabled}
                  setActiveTab={setActiveTab}
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
                  context={context}
                />
              </div>
            )}
          </div>
        </div>
      </SharePointDialog>
    </>
  );
};

export default AlertSettingsTabs;
