import * as React from "react";
import { Settings24Regular, Add24Regular, Delete24Regular, Save24Regular, Dismiss24Regular } from "@fluentui/react-icons";
import SharePointDialog from "../UI/SharePointDialog";
import {
  SharePointButton,
  SharePointInput,
  SharePointTextArea,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
  ISharePointSelectOption
} from "../UI/SharePointControls";
import { AlertPriority, NotificationType, IAlertItem, IAlertType } from "../Alerts/IAlerts";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from "./ComprehensiveAlertSettings.module.scss";

export interface IComprehensiveAlertSettingsProps {
  isInEditMode: boolean;
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
  graphClient: MSGraphClientV3;
  onSettingsChange: (settings: ISettingsData) => void;
}

export interface ISettingsData {
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
}

interface INewAlert {
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl: string;
  linkDescription: string;
}

const ComprehensiveAlertSettings: React.FC<IComprehensiveAlertSettingsProps> = ({
  isInEditMode,
  alertTypesJson,
  userTargetingEnabled,
  notificationsEnabled,
  richMediaEnabled,
  graphClient,
  onSettingsChange
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<"settings" | "alerts" | "types">("settings");
  const [isCreatingAlert, setIsCreatingAlert] = React.useState(false);
  const [isCreatingType, setIsCreatingType] = React.useState(false);

  // Settings state
  const [settings, setSettings] = React.useState<ISettingsData>({
    alertTypesJson,
    userTargetingEnabled,
    notificationsEnabled,
    richMediaEnabled
  });

  // Alert types state
  const [alertTypes, setAlertTypes] = React.useState<IAlertType[]>(() => {
    try {
      return JSON.parse(alertTypesJson);
    } catch {
      return [];
    }
  });

  // New alert state
  const [newAlert, setNewAlert] = React.useState<INewAlert>({
    title: "",
    description: "",
    AlertType: "",
    priority: AlertPriority.Medium,
    isPinned: false,
    notificationType: NotificationType.None,
    linkUrl: "",
    linkDescription: ""
  });

  // New alert type state
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
      [AlertPriority.Low]: ""
    }
  });

  // Don't render if not in edit mode
  if (!isInEditMode) {
    return null;
  }

  // Priority options
  const priorityOptions: ISharePointSelectOption[] = [
    { value: AlertPriority.Low, label: "Low" },
    { value: AlertPriority.Medium, label: "Medium" },
    { value: AlertPriority.High, label: "High" },
    { value: AlertPriority.Critical, label: "Critical" }
  ];

  // Notification type options
  const notificationOptions: ISharePointSelectOption[] = [
    { value: NotificationType.None, label: "None" },
    { value: NotificationType.Browser, label: "Browser Notification" },
    { value: NotificationType.Email, label: "Email Notification" },
    { value: NotificationType.Both, label: "Both Browser and Email" }
  ];

  // Alert type options
  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(type => ({
    value: type.name,
    label: type.name
  }));

  const handleSaveSettings = () => {
    const updatedSettings = {
      ...settings,
      alertTypesJson: JSON.stringify(alertTypes, null, 2)
    };
    onSettingsChange(updatedSettings);
    setIsOpen(false);
  };

  const handleCreateAlert = async () => {
    if (!newAlert.title || !newAlert.description || !newAlert.AlertType) {
      alert("Please fill in all required fields.");
      return;
    }

    try {
      // Create the alert item
      const alertItem: Partial<IAlertItem> = {
        title: newAlert.title,
        description: newAlert.description,
        AlertType: newAlert.AlertType,
        priority: newAlert.priority,
        isPinned: newAlert.isPinned,
        notificationType: newAlert.notificationType,
        ...(newAlert.linkUrl && newAlert.linkDescription && {
          link: {
            Url: newAlert.linkUrl,
            Description: newAlert.linkDescription
          }
        }),
        createdDate: new Date().toISOString(),
        createdBy: "Alert Settings"
      };

      // TODO: Here you would save the alert to SharePoint
      // For now, we'll just show a success message
      console.log("Creating alert:", alertItem);
      alert("Alert created successfully! (In a real implementation, this would save to SharePoint)");

      // Reset form
      setNewAlert({
        title: "",
        description: "",
        AlertType: "",
        priority: AlertPriority.Medium,
        isPinned: false,
        notificationType: NotificationType.None,
        linkUrl: "",
        linkDescription: ""
      });
      setIsCreatingAlert(false);

    } catch (error) {
      console.error("Error creating alert:", error);
      alert("Failed to create alert. Please try again.");
    }
  };

  const handleCreateAlertType = () => {
    if (!newAlertType.name.trim()) {
      alert("Please enter an alert type name.");
      return;
    }

    const updatedTypes = [...alertTypes, { ...newAlertType, name: newAlertType.name.trim() }];
    setAlertTypes(updatedTypes);

    // Reset form
    setNewAlertType({
      name: "",
      iconName: "Info",
      backgroundColor: "#0078d4",
      textColor: "#ffffff",
      additionalStyles: "",
      priorityStyles: {
        [AlertPriority.Critical]: "border: 2px solid #E81123;",
        [AlertPriority.High]: "border: 1px solid #EA4300;",
        [AlertPriority.Medium]: "",
        [AlertPriority.Low]: ""
      }
    });
    setIsCreatingType(false);
  };

  const handleDeleteAlertType = (index: number) => {
    if (confirm("Are you sure you want to delete this alert type?")) {
      const updatedTypes = alertTypes.filter((_, i) => i !== index);
      setAlertTypes(updatedTypes);
    }
  };

  const renderSettings = () => (
    <div className={styles.tabContent}>
      <SharePointSection title="Feature Settings">
        <SharePointToggle
          label="Enable User Targeting"
          checked={settings.userTargetingEnabled}
          onChange={(checked) => setSettings(prev => ({ ...prev, userTargetingEnabled: checked }))}
          description="Allow alerts to target specific users or groups based on SharePoint profiles"
        />

        <SharePointToggle
          label="Enable Browser Notifications"
          checked={settings.notificationsEnabled}
          onChange={(checked) => setSettings(prev => ({ ...prev, notificationsEnabled: checked }))}
          description="Send browser notifications for critical and high-priority alerts"
        />

        <SharePointToggle
          label="Enable Rich Media Support"
          checked={settings.richMediaEnabled}
          onChange={(checked) => setSettings(prev => ({ ...prev, richMediaEnabled: checked }))}
          description="Support images, videos, HTML content, and markdown in alert descriptions"
        />
      </SharePointSection>

      <SharePointSection title="Advanced Configuration">
        <div className={styles.infoBox}>
          <h4>Alert Types JSON Configuration</h4>
          <p>Advanced users can modify the alert types directly in JSON format. This allows for complete customization of colors, styles, and priority configurations.</p>
        </div>

        <SharePointTextArea
          label="Alert Types Configuration"
          value={JSON.stringify(alertTypes, null, 2)}
          onChange={(value) => {
            try {
              const parsed = JSON.parse(value);
              setAlertTypes(parsed);
            } catch {
              // Invalid JSON, don't update
            }
          }}
          rows={10}
          description="Modify the JSON configuration for alert types. Use the 'Alert Types' tab for a user-friendly editor."
          className={styles.jsonEditor}
        />
      </SharePointSection>
    </div>
  );

  const renderAlerts = () => (
    <div className={styles.tabContent}>
      <div className={styles.tabHeader}>
        <h3>Create New Alert</h3>
        <SharePointButton
          variant="primary"
          icon={<Add24Regular />}
          onClick={() => setIsCreatingAlert(true)}
        >
          New Alert
        </SharePointButton>
      </div>

      {isCreatingAlert && (
        <SharePointSection title="Create Alert">
          <div className={styles.formGrid}>
            <SharePointInput
              label="Alert Title"
              value={newAlert.title}
              onChange={(value) => setNewAlert(prev => ({ ...prev, title: value }))}
              placeholder="Enter alert title"
              required
            />

            <SharePointSelect
              label="Alert Type"
              value={newAlert.AlertType}
              onChange={(value) => setNewAlert(prev => ({ ...prev, AlertType: value }))}
              options={alertTypeOptions}
              placeholder="Select alert type"
              required
            />

            <SharePointSelect
              label="Priority"
              value={newAlert.priority}
              onChange={(value) => setNewAlert(prev => ({ ...prev, priority: value as AlertPriority }))}
              options={priorityOptions}
            />

            <SharePointSelect
              label="Notification Type"
              value={newAlert.notificationType}
              onChange={(value) => setNewAlert(prev => ({ ...prev, notificationType: value as NotificationType }))}
              options={notificationOptions}
            />
          </div>

          <SharePointTextArea
            label="Alert Description"
            value={newAlert.description}
            onChange={(value) => setNewAlert(prev => ({ ...prev, description: value }))}
            placeholder="Enter alert description (supports HTML and markdown)"
            required
            rows={4}
          />

          <div className={styles.formGrid}>
            <SharePointInput
              label="Link URL"
              value={newAlert.linkUrl}
              onChange={(value) => setNewAlert(prev => ({ ...prev, linkUrl: value }))}
              placeholder="https://example.com"
              type="url"
            />

            <SharePointInput
              label="Link Description"
              value={newAlert.linkDescription}
              onChange={(value) => setNewAlert(prev => ({ ...prev, linkDescription: value }))}
              placeholder="Learn more"
            />
          </div>

          <SharePointToggle
            label="Pin Alert"
            checked={newAlert.isPinned}
            onChange={(checked) => setNewAlert(prev => ({ ...prev, isPinned: checked }))}
            description="Pinned alerts stay at the top and remain visible across page navigation"
          />

          <div className={styles.formActions}>
            <SharePointButton
              variant="primary"
              icon={<Save24Regular />}
              onClick={handleCreateAlert}
            >
              Create Alert
            </SharePointButton>
            <SharePointButton
              variant="secondary"
              icon={<Dismiss24Regular />}
              onClick={() => setIsCreatingAlert(false)}
            >
              Cancel
            </SharePointButton>
          </div>
        </SharePointSection>
      )}

      <div className={styles.infoBox}>
        <h4>Alert Management</h4>
        <p>Create alerts that will appear in the banner for all users (or targeted users if targeting is enabled).
          Alerts support rich content, links, priority levels, and various notification options.</p>
        <ul>
          <li><strong>Critical:</strong> Red styling, optional notifications</li>
          <li><strong>High:</strong> Orange styling, recommended for important updates</li>
          <li><strong>Medium:</strong> Blue styling, for general information</li>
          <li><strong>Low:</strong> Green styling, for tips and suggestions</li>
        </ul>
      </div>
    </div>
  );

  const renderAlertTypes = () => (
    <div className={styles.tabContent}>
      <div className={styles.tabHeader}>
        <h3>Alert Types</h3>
        <SharePointButton
          variant="primary"
          icon={<Add24Regular />}
          onClick={() => setIsCreatingType(true)}
        >
          New Alert Type
        </SharePointButton>
      </div>

      {isCreatingType && (
        <SharePointSection title="Create Alert Type">
          <div className={styles.formGrid}>
            <SharePointInput
              label="Type Name"
              value={newAlertType.name}
              onChange={(value) => setNewAlertType(prev => ({ ...prev, name: value }))}
              placeholder="e.g., Maintenance, Emergency, Update"
              required
            />

            <SharePointInput
              label="Icon Name"
              value={newAlertType.iconName}
              onChange={(value) => setNewAlertType(prev => ({ ...prev, iconName: value }))}
              placeholder="Info, Warning, Error, etc."
            />

            <SharePointInput
              label="Background Color"
              value={newAlertType.backgroundColor}
              onChange={(value) => setNewAlertType(prev => ({ ...prev, backgroundColor: value }))}
              placeholder="#0078d4"
              type="text"
            />

            <SharePointInput
              label="Text Color"
              value={newAlertType.textColor}
              onChange={(value) => setNewAlertType(prev => ({ ...prev, textColor: value }))}
              placeholder="#ffffff"
              type="text"
            />
          </div>

          <SharePointTextArea
            label="Additional Styles"
            value={newAlertType.additionalStyles || ""}
            onChange={(value) => setNewAlertType(prev => ({ ...prev, additionalStyles: value }))}
            placeholder="Custom CSS styles"
            rows={3}
          />

          <div className={styles.formActions}>
            <SharePointButton
              variant="primary"
              icon={<Save24Regular />}
              onClick={handleCreateAlertType}
            >
              Create Type
            </SharePointButton>
            <SharePointButton
              variant="secondary"
              icon={<Dismiss24Regular />}
              onClick={() => setIsCreatingType(false)}
            >
              Cancel
            </SharePointButton>
          </div>
        </SharePointSection>
      )}

      <SharePointSection title="Existing Alert Types">
        <div className={styles.alertTypesList}>
          {alertTypes.map((type, index) => (
            <div key={type.name} className={styles.alertTypeItem}>
              <div className={styles.alertTypePreview} style={{ backgroundColor: type.backgroundColor, color: type.textColor }}>
                {type.name}
              </div>
              <div className={styles.alertTypeInfo}>
                <strong>{type.name}</strong>
                <span>Icon: {type.iconName}</span>
                <span>Colors: {type.backgroundColor} / {type.textColor}</span>
              </div>
              <SharePointButton
                variant="danger"
                icon={<Delete24Regular />}
                onClick={() => handleDeleteAlertType(index)}
              >
                Delete
              </SharePointButton>
            </div>
          ))}
        </div>
      </SharePointSection>
    </div>
  );

  const footer = (
    <div className={styles.dialogFooter}>
      <SharePointButton variant="primary" onClick={handleSaveSettings}>
        Save Settings
      </SharePointButton>
      <SharePointButton variant="secondary" onClick={() => setIsOpen(false)}>
        Cancel
      </SharePointButton>
    </div>
  );

  return (
    <>
      <SharePointButton
        variant="secondary"
        icon={<Settings24Regular />}
        onClick={() => setIsOpen(true)}
        className={styles.settingsButton}
      >
        Settings
      </SharePointButton>

      <SharePointDialog
        isOpen={isOpen}
        onClose={() => setIsOpen(false)}
        title="Alert Banner Settings"
        width={900}
        height={700}
        footer={footer}
      >
        <div className={styles.settingsContainer}>
          <div className={styles.tabs}>
            <button
              className={`${styles.tab} ${activeTab === "settings" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("settings")}
            >
              Settings
            </button>
            <button
              className={`${styles.tab} ${activeTab === "alerts" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("alerts")}
            >
              Create Alerts
            </button>
            <button
              className={`${styles.tab} ${activeTab === "types" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("types")}
            >
              Alert Types
            </button>
          </div>

          {activeTab === "settings" && renderSettings()}
          {activeTab === "alerts" && renderAlerts()}
          {activeTab === "types" && renderAlertTypes()}
        </div>
      </SharePointDialog>
    </>
  );
};

export default ComprehensiveAlertSettings;