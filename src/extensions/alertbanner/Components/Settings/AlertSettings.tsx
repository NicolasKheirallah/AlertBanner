import * as React from "react";
import { Settings24Regular, Add24Regular, Delete24Regular, Save24Regular, Dismiss24Regular, Eye24Regular } from "@fluentui/react-icons";
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
import ColorPicker from "../UI/ColorPicker";
import AlertPreview from "../UI/AlertPreview";
import AlertTemplates, { IAlertTemplate } from "../UI/AlertTemplates";
import { AlertPriority, NotificationType, IAlertType } from "../Alerts/IAlerts";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import SharePointRichTextEditor from "../UI/SharePointRichTextEditor";
import SiteSelector from "../UI/SiteSelector";
import { SiteContextDetector, ISiteContext, ISiteValidationResult } from "../Utils/SiteContextDetector";
import { SharePointAlertService, IAlertItem } from "../Services/SharePointAlertService";
import LanguageFieldManager from "../UI/LanguageFieldManager";
import styles from "./AlertSettings.module.scss";

export interface IAlertSettingsProps {
  isInEditMode: boolean;
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
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
  targetSites: string[];
  scheduledStart?: Date;
  scheduledEnd?: Date;
}

interface IFormErrors {
  title?: string;
  description?: string;
  AlertType?: string;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
}

interface IEditingAlert extends Omit<IAlertItem, 'scheduledStart' | 'scheduledEnd'> {
  scheduledStart?: Date;
  scheduledEnd?: Date;
}

const AlertSettings: React.FC<IAlertSettingsProps> = ({
  isInEditMode,
  alertTypesJson,
  userTargetingEnabled,
  notificationsEnabled,
  richMediaEnabled,
  graphClient,
  context,
  onSettingsChange
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [activeTab, setActiveTab] = React.useState<"create" | "manage" | "types" | "settings">("create");
  const [showTemplates, setShowTemplates] = React.useState(true);
  const [isCreatingType, setIsCreatingType] = React.useState(false);
  const [showPreview, setShowPreview] = React.useState(true);

  // Site context and targeting
  const [siteDetector] = React.useState(() => new SiteContextDetector(graphClient, context));
  const [currentSiteContext, setCurrentSiteContext] = React.useState<ISiteContext | null>(null);
  const [creationProgress, setCreationProgress] = React.useState<ISiteValidationResult[]>([]);
  const [isCreatingAlert, setIsCreatingAlert] = React.useState(false);

  // SharePoint service
  const [alertService] = React.useState(() => new SharePointAlertService(graphClient, context));

  // Alert management state
  const [existingAlerts, setExistingAlerts] = React.useState<IAlertItem[]>([]);
  const [isLoadingAlerts, setIsLoadingAlerts] = React.useState(false);
  const [isCreatingLists, setIsCreatingLists] = React.useState(false);
  const [selectedAlerts, setSelectedAlerts] = React.useState<string[]>([]);
  const [editingAlert, setEditingAlert] = React.useState<IEditingAlert | null>(null);
  const [isEditingAlert, setIsEditingAlert] = React.useState(false);

  // Settings state
  const [settings, setSettings] = React.useState<ISettingsData>({
    alertTypesJson,
    userTargetingEnabled,
    notificationsEnabled,
    richMediaEnabled
  });

  // Alert types state - load from SharePoint instead of JSON
  const [alertTypes, setAlertTypes] = React.useState<IAlertType[]>([]);

  // Load alert types from SharePoint on init
  React.useEffect(() => {
    if (isInEditMode) {
      alertService.getAlertTypes().then(types => {
        if (types.length > 0) {
          setAlertTypes(types);
        }
      }).catch(error => {
        console.warn('Using default alert types:', error);
      });
    }
  }, [isInEditMode, alertService]);

  // New alert state
  const [newAlert, setNewAlert] = React.useState<INewAlert>({
    title: "",
    description: "",
    AlertType: "",
    priority: AlertPriority.Medium,
    isPinned: false,
    notificationType: NotificationType.None,
    linkUrl: "",
    linkDescription: "",
    targetSites: [],
    scheduledStart: undefined,
    scheduledEnd: undefined
  });

  // Form validation
  const [errors, setErrors] = React.useState<IFormErrors>({});

  // New alert type state with better defaults
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

  // Initialize site context
  React.useEffect(() => {
    if (isInEditMode) {
      siteDetector.getCurrentSiteContext().then(context => {
        setCurrentSiteContext(context);
        // Set current site as default target if no sites selected
        if (newAlert.targetSites.length === 0) {
          setNewAlert(prev => ({
            ...prev,
            targetSites: [context.siteId]
          }));
        }
      }).catch(error => {
        console.error('Failed to get site context:', error);
      });
    }
  }, [isInEditMode, siteDetector]);

  // Don't render if not in edit mode
  if (!isInEditMode) {
    return null;
  }

  // Priority options
  const priorityOptions: ISharePointSelectOption[] = [
    { value: AlertPriority.Low, label: "Low Priority - Informational updates" },
    { value: AlertPriority.Medium, label: "Medium Priority - General announcements" },
    { value: AlertPriority.High, label: "High Priority - Important updates" },
    { value: AlertPriority.Critical, label: "Critical Priority - Urgent action required" }
  ];

  // Notification type options
  const notificationOptions: ISharePointSelectOption[] = [
    { value: NotificationType.None, label: "No notifications" },
    { value: NotificationType.Browser, label: "Browser notification only" },
    { value: NotificationType.Email, label: "Email notification only" },
    { value: NotificationType.Both, label: "Both browser and email" }
  ];

  // Alert type options
  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(type => ({
    value: type.name,
    label: `${type.name} (${type.backgroundColor})`
  }));

  // Validation functions
  const validateAlert = (): boolean => {
    const newErrors: IFormErrors = {};

    if (!newAlert.title.trim()) {
      newErrors.title = "Alert title is required";
    } else if (newAlert.title.length < 3) {
      newErrors.title = "Title must be at least 3 characters";
    } else if (newAlert.title.length > 100) {
      newErrors.title = "Title must be less than 100 characters";
    }

    if (!newAlert.description.trim()) {
      newErrors.description = "Alert description is required";
    } else if (newAlert.description.replace(/<[^>]*>/g, '').length < 10) {
      newErrors.description = "Description must be at least 10 characters (excluding HTML tags)";
    }

    if (!newAlert.AlertType) {
      newErrors.AlertType = "Please select an alert type";
    }

    if (newAlert.targetSites.length === 0) {
      newErrors.targetSites = "Please select at least one site for alert distribution";
    }

    if (newAlert.linkUrl && !isValidUrl(newAlert.linkUrl)) {
      newErrors.linkUrl = "Please enter a valid URL";
    }

    if (newAlert.linkUrl && !newAlert.linkDescription.trim()) {
      newErrors.linkDescription = "Link description is required when URL is provided";
    }

    if (newAlert.scheduledStart && newAlert.scheduledEnd) {
      if (newAlert.scheduledEnd <= newAlert.scheduledStart) {
        newErrors.scheduledEnd = "End date must be after start date";
      }
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const isValidUrl = (url: string): boolean => {
    try {
      new URL(url);
      return true;
    } catch {
      return false;
    }
  };

  const handleTemplateSelect = (template: IAlertTemplate) => {
    setNewAlert(prev => ({
      title: template.template.title,
      description: template.template.description,
      AlertType: alertTypes.length > 0 ? alertTypes[0].name : "",
      priority: template.template.priority,
      isPinned: template.template.isPinned,
      notificationType: template.template.notificationType,
      linkUrl: template.template.linkUrl || "",
      linkDescription: template.template.linkDescription || "",
      // Keep existing targeting settings
      targetSites: prev.targetSites.length > 0 ? prev.targetSites : (currentSiteContext ? [currentSiteContext.siteId] : []),
      scheduledStart: prev.scheduledStart,
      scheduledEnd: prev.scheduledEnd
    }));
    setShowTemplates(false);
    setErrors({});
  };

  const handleSaveSettings = async () => {
    try {
      // Save alert types to SharePoint
      await alertService.saveAlertTypes(alertTypes);

      const updatedSettings = {
        ...settings,
        alertTypesJson: JSON.stringify(alertTypes, null, 2)
      };
      onSettingsChange(updatedSettings);
      setIsOpen(false);

      // Show success message
      const successElement = document.createElement('div');
      successElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #107c10;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
      `;
      successElement.textContent = '✅ Settings saved successfully!';
      document.body.appendChild(successElement);

      setTimeout(() => {
        if (document.body.contains(successElement)) {
          document.body.removeChild(successElement);
        }
      }, 3000);
    } catch (error) {
      console.error('Failed to save settings:', error);

      // Still save to local settings even if SharePoint fails
      const updatedSettings = {
        ...settings,
        alertTypesJson: JSON.stringify(alertTypes, null, 2)
      };
      onSettingsChange(updatedSettings);
      setIsOpen(false);

      // Show appropriate warning message based on error type
      const warningElement = document.createElement('div');
      let message = '';
      let backgroundColor = '#8a6914';

      if (error.message?.includes('PERMISSION_DENIED')) {
        message = '⚠️ Settings saved locally only - SharePoint permissions required for persistent storage';
      } else if (error.message?.includes('LISTS_NOT_FOUND')) {
        message = '⚠️ Settings saved locally only - SharePoint lists not available';
      } else {
        message = '⚠️ Settings saved locally only - SharePoint integration unavailable';
      }

      warningElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${backgroundColor};
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 320px;
        line-height: 1.4;
      `;
      warningElement.textContent = message;
      document.body.appendChild(warningElement);

      setTimeout(() => {
        if (document.body.contains(warningElement)) {
          document.body.removeChild(warningElement);
        }
      }, 8000);
    }
  };

  const handleCreateAlert = async () => {
    if (!validateAlert()) {
      return;
    }

    setIsCreatingAlert(true);
    setCreationProgress([]);

    try {
      // Validate site permissions first
      const siteValidations = await siteDetector.validateSiteAccess(newAlert.targetSites);
      setCreationProgress(siteValidations);

      // Filter sites where user can create alerts
      const validSites = siteValidations.filter(s => s.canCreateAlerts);

      if (validSites.length === 0) {
        throw new Error("You don't have permission to create alerts on any of the selected sites.");
      }

      // Create alert using SharePoint service
      const alertItem = {
        title: newAlert.title.trim(),
        description: newAlert.description.trim(),
        AlertType: newAlert.AlertType,
        priority: newAlert.priority,
        isPinned: newAlert.isPinned,
        notificationType: newAlert.notificationType,
        linkUrl: newAlert.linkUrl?.trim(),
        linkDescription: newAlert.linkDescription?.trim(),
        targetSites: newAlert.targetSites,
        scheduledStart: newAlert.scheduledStart?.toISOString(),
        scheduledEnd: newAlert.scheduledEnd?.toISOString(),
        metadata: {
          sourceSiteId: currentSiteContext?.siteId,
          sourceSiteName: currentSiteContext?.siteName,
          deploymentTargets: newAlert.targetSites
        }
      };

      // Create the alert in SharePoint
      const createdAlert = await alertService.createAlert(alertItem);

      // Add to local state
      setExistingAlerts(prev => [createdAlert, ...prev]);

      // Show success message
      const alertElement = document.createElement('div');
      alertElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #107c10;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 300px;
      `;
      alertElement.textContent = `✅ Alert created successfully!`;
      document.body.appendChild(alertElement);

      setTimeout(() => {
        if (document.body.contains(alertElement)) {
          document.body.removeChild(alertElement);
        }
      }, 3000);

      // Reset form
      resetForm();

    } catch (error) {
      console.error("Error creating alert:", error);
      const alertElement = document.createElement('div');
      alertElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #d13438;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
      `;
      alertElement.textContent = `❌ Failed to create alert: ${error.message}`;
      document.body.appendChild(alertElement);

      setTimeout(() => {
        if (document.body.contains(alertElement)) {
          document.body.removeChild(alertElement);
        }
      }, 5000);
    } finally {
      setIsCreatingAlert(false);
    }
  };

  const resetForm = () => {
    setNewAlert({
      title: "",
      description: "",
      AlertType: "",
      priority: AlertPriority.Medium,
      isPinned: false,
      notificationType: NotificationType.None,
      linkUrl: "",
      linkDescription: "",
      targetSites: currentSiteContext ? [currentSiteContext.siteId] : [],
      scheduledStart: undefined,
      scheduledEnd: undefined
    });
    setErrors({});
    setShowTemplates(true);
    setCreationProgress([]);
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

    // Show success message
    const successElement = document.createElement('div');
    successElement.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #107c10;
      color: white;
      padding: 16px 20px;
      border-radius: 4px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
      z-index: 10001;
      font-family: Segoe UI;
      font-size: 14px;
    `;
    successElement.textContent = "✅ Alert type created successfully!";
    document.body.appendChild(successElement);

    setTimeout(() => {
      if (document.body.contains(successElement)) {
        document.body.removeChild(successElement);
      }
    }, 3000);
  };

  const handleDeleteAlertType = (index: number) => {
    const typeToDelete = alertTypes[index];
    if (confirm(`Are you sure you want to delete the "${typeToDelete.name}" alert type? This action cannot be undone.`)) {
      const updatedTypes = alertTypes.filter((_, i) => i !== index);
      setAlertTypes(updatedTypes);

      // If the deleted type was selected in the new alert, clear it
      if (newAlert.AlertType === typeToDelete.name) {
        setNewAlert(prev => ({ ...prev, AlertType: "" }));
      }
    }
  };

  // Drag and drop functionality for alert types
  const handleDragStart = (e: React.DragEvent<HTMLDivElement>, index: number) => {
    e.dataTransfer.setData('text/plain', index.toString());
    e.currentTarget.style.opacity = '0.5';
  };

  const handleDragEnd = (e: React.DragEvent<HTMLDivElement>) => {
    e.currentTarget.style.opacity = '1';
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>, dropIndex: number) => {
    e.preventDefault();
    const dragIndex = parseInt(e.dataTransfer.getData('text/plain'), 10);

    if (dragIndex === dropIndex) return;

    const newTypes = [...alertTypes];
    const [draggedItem] = newTypes.splice(dragIndex, 1);
    newTypes.splice(dropIndex, 0, draggedItem);

    setAlertTypes(newTypes);

    // Show success message
    const successElement = document.createElement('div');
    successElement.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #107c10;
      color: white;
      padding: 16px 20px;
      border-radius: 4px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
      z-index: 10001;
      font-family: Segoe UI;
      font-size: 14px;
    `;
    successElement.textContent = '✅ Alert types reordered successfully!';
    document.body.appendChild(successElement);

    setTimeout(() => {
      if (document.body.contains(successElement)) {
        document.body.removeChild(successElement);
      }
    }, 2000);
  };

  // Load existing alerts
  const loadExistingAlerts = async () => {
    setIsLoadingAlerts(true);
    try {
      // Initialize SharePoint lists if needed
      await alertService.initializeLists();

      // Load alerts from SharePoint
      const alerts = await alertService.getAlerts();
      setExistingAlerts(alerts);

      // If no alerts were loaded and SharePoint integration is working, 
      // show a helpful message about sample data being created
      if (alerts.length === 0) {
        console.log('No alerts found. Sample alerts should have been created during initialization.');

        // Try loading again after a brief delay to account for list creation timing
        setTimeout(async () => {
          try {
            const retryAlerts = await alertService.getAlerts();
            if (retryAlerts.length > 0) {
              setExistingAlerts(retryAlerts);
            }
          } catch (retryError) {
            console.warn('Retry loading alerts failed:', retryError);
          }
        }, 2000);
      }
    } catch (error) {
      console.error('Failed to load alerts:', error);

      // Set empty alerts list when SharePoint integration fails
      setExistingAlerts([]);

      // Show appropriate error message based on error type
      const errorElement = document.createElement('div');
      let message = '';
      let backgroundColor = '#d13438';

      if (error.message?.includes('PERMISSION_DENIED')) {
        backgroundColor = '#8a6914';
        message = `
          <strong>⚠️ Limited Permissions</strong><br>
          You don't have permission to create SharePoint lists. Please contact your administrator to set up the required lists.
        `;
      } else if (error.message?.includes('LISTS_NOT_FOUND')) {
        backgroundColor = '#8a6914';
        message = `
          <strong>⚠️ Lists Not Found</strong><br>
          SharePoint lists don't exist and cannot be created. Please contact your administrator to set up the Alert Banner lists.
        `;
      } else {
        message = `
          <strong>❌ SharePoint Error</strong><br>
          Unable to connect to SharePoint. Please check your permissions and try again.
        `;
      }

      errorElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${backgroundColor};
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 320px;
        line-height: 1.4;
      `;
      errorElement.innerHTML = message;
      document.body.appendChild(errorElement);

      setTimeout(() => {
        if (document.body.contains(errorElement)) {
          document.body.removeChild(errorElement);
        }
      }, 8000);
    } finally {
      setIsLoadingAlerts(false);
    }
  };

  const handleCreateLists = async () => {
    if (!confirm('This will create the Alert Banner lists (Alerts and AlertBannerTypes) on the current site. Continue?')) {
      return;
    }

    setIsCreatingLists(true);
    try {
      // Initialize/create SharePoint lists
      await alertService.initializeLists();
      
      // Show success message
      const successElement = document.createElement('div');
      successElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #107c10;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 320px;
        line-height: 1.4;
      `;
      successElement.textContent = '✅ Alert Banner lists created successfully! You can now create and manage alerts.';
      document.body.appendChild(successElement);

      setTimeout(() => {
        if (document.body.contains(successElement)) {
          document.body.removeChild(successElement);
        }
      }, 5000);

      // Automatically refresh the alerts after successful list creation
      setTimeout(() => {
        loadExistingAlerts();
      }, 1000);

    } catch (error) {
      console.error('Failed to create lists:', error);
      
      // Show error message based on error type
      const errorElement = document.createElement('div');
      let message = '';
      let backgroundColor = '#d13438';

      if (error.message?.includes('PERMISSION_DENIED') || error.message?.includes('403')) {
        backgroundColor = '#8a6914';
        message = '⚠️ Insufficient permissions to create SharePoint lists. Please contact your site administrator.';
      } else if (error.message?.includes('ALREADY_EXISTS')) {
        backgroundColor = '#107c10';
        message = '✅ Lists already exist! Refreshing content...';
        // If lists already exist, just refresh
        setTimeout(() => loadExistingAlerts(), 500);
      } else {
        message = '❌ Failed to create Alert Banner lists. Please check your permissions and try again.';
      }

      errorElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${backgroundColor};
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 320px;
        line-height: 1.4;
      `;
      errorElement.textContent = message;
      document.body.appendChild(errorElement);

      setTimeout(() => {
        if (document.body.contains(errorElement)) {
          document.body.removeChild(errorElement);
        }
      }, 8000);
    } finally {
      setIsCreatingLists(false);
    }
  };

  const handleDeleteAlert = async (alertId: string) => {
    if (confirm('Are you sure you want to delete this alert? This action cannot be undone.')) {
      try {
        // Delete from SharePoint
        await alertService.deleteAlert(alertId);

        // Update local state
        setExistingAlerts(prev => prev.filter(alert => alert.id !== alertId));
        setSelectedAlerts(prev => prev.filter(id => id !== alertId));

        // Show success message
        const successElement = document.createElement('div');
        successElement.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          background: #107c10;
          color: white;
          padding: 16px 20px;
          border-radius: 4px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.2);
          z-index: 10001;
          font-family: Segoe UI;
          font-size: 14px;
        `;
        successElement.textContent = '✅ Alert deleted successfully!';
        document.body.appendChild(successElement);

        setTimeout(() => {
          if (document.body.contains(successElement)) {
            document.body.removeChild(successElement);
          }
        }, 3000);
      } catch (error) {
        console.error('Failed to delete alert:', error);

        // Show error message
        const errorElement = document.createElement('div');
        errorElement.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          background: #d13438;
          color: white;
          padding: 16px 20px;
          border-radius: 4px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.2);
          z-index: 10001;
          font-family: Segoe UI;
          font-size: 14px;
        `;
        errorElement.textContent = '❌ Failed to delete alert';
        document.body.appendChild(errorElement);

        setTimeout(() => {
          if (document.body.contains(errorElement)) {
            document.body.removeChild(errorElement);
          }
        }, 3000);
      }
    }
  };

  const handleBulkDelete = async () => {
    if (selectedAlerts.length === 0) return;

    if (confirm(`Are you sure you want to delete ${selectedAlerts.length} alert(s)? This action cannot be undone.`)) {
      try {
        // Delete from SharePoint
        await alertService.deleteAlerts(selectedAlerts);

        // Update local state
        setExistingAlerts(prev => prev.filter(alert => !selectedAlerts.includes(alert.id)));
        setSelectedAlerts([]);

        // Show success message
        const successElement = document.createElement('div');
        successElement.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          background: #107c10;
          color: white;
          padding: 16px 20px;
          border-radius: 4px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.2);
          z-index: 10001;
          font-family: Segoe UI;
          font-size: 14px;
        `;
        successElement.textContent = `✅ ${selectedAlerts.length} alert(s) deleted successfully!`;
        document.body.appendChild(successElement);

        setTimeout(() => {
          if (document.body.contains(successElement)) {
            document.body.removeChild(successElement);
          }
        }, 3000);
      } catch (error) {
        console.error('Failed to delete alerts:', error);

        // Show error message
        const errorElement = document.createElement('div');
        errorElement.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          background: #d13438;
          color: white;
          padding: 16px 20px;
          border-radius: 4px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.2);
          z-index: 10001;
          font-family: Segoe UI;
          font-size: 14px;
        `;
        errorElement.textContent = '❌ Failed to delete some alerts';
        document.body.appendChild(errorElement);

        setTimeout(() => {
          if (document.body.contains(errorElement)) {
            document.body.removeChild(errorElement);
          }
        }, 3000);
      }
    }
  };

  const handleEditAlert = (alert: IAlertItem) => {
    setEditingAlert({
      ...alert,
      scheduledStart: alert.scheduledStart ? new Date(alert.scheduledStart) : undefined,
      scheduledEnd: alert.scheduledEnd ? new Date(alert.scheduledEnd) : undefined
    });
    setIsEditingAlert(true);
  };

  const handleSaveEditedAlert = async () => {
    if (!editingAlert) return;

    try {
      // Update in SharePoint
      const updatedAlert = await alertService.updateAlert(editingAlert.id, {
        ...editingAlert,
        scheduledStart: editingAlert.scheduledStart?.toISOString(),
        scheduledEnd: editingAlert.scheduledEnd?.toISOString()
      });

      // Update local state
      setExistingAlerts(prev => prev.map(alert =>
        alert.id === editingAlert.id ? updatedAlert : alert
      ));

      setIsEditingAlert(false);
      setEditingAlert(null);

      // Show success message
      const successElement = document.createElement('div');
      successElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #107c10;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
      `;
      successElement.textContent = '✅ Alert updated successfully!';
      document.body.appendChild(successElement);

      setTimeout(() => {
        if (document.body.contains(successElement)) {
          document.body.removeChild(successElement);
        }
      }, 3000);
    } catch (error) {
      console.error('Failed to update alert:', error);

      // Show error message
      const errorElement = document.createElement('div');
      errorElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #d13438;
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
      `;
      errorElement.textContent = '❌ Failed to update alert';
      document.body.appendChild(errorElement);

      setTimeout(() => {
        if (document.body.contains(errorElement)) {
          document.body.removeChild(errorElement);
        }
      }, 3000);
    }
  };

  const handleCancelEdit = () => {
    setIsEditingAlert(false);
    setEditingAlert(null);
  };

  // Load alerts when management tab is opened
  React.useEffect(() => {
    if (activeTab === 'manage' && existingAlerts.length === 0 && !isLoadingAlerts) {
      loadExistingAlerts();
    }
  }, [activeTab]);

  const getSelectedAlertType = (): IAlertType | undefined => {
    return alertTypes.find(type => type.name === newAlert.AlertType);
  };

  const renderCreateAlert = () => (
    <div className={styles.tabContent}>
      {showTemplates && (
        <div className={styles.templatesSection}>
          <AlertTemplates
            onSelectTemplate={handleTemplateSelect}
            className={styles.templates}
          />

          <div className={styles.templateActions}>
            <SharePointButton
              variant="secondary"
              onClick={() => setShowTemplates(false)}
            >
              Start from Scratch
            </SharePointButton>
          </div>
        </div>
      )}

      {!showTemplates && (
        <div className={styles.alertForm}>
          <div className={styles.formWithPreview}>
            <div className={styles.formColumn}>
              <SharePointSection title="Basic Information">
                <SharePointInput
                  label="Alert Title"
                  value={newAlert.title}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, title: value }));
                    if (errors.title) setErrors(prev => ({ ...prev, title: undefined }));
                  }}
                  placeholder="Enter a clear, concise title"
                  required
                  error={errors.title}
                  description="This will be the main heading of your alert (3-100 characters)"
                />

                <SharePointRichTextEditor
                  label="Alert Description"
                  value={newAlert.description}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, description: value }));
                    if (errors.description) setErrors(prev => ({ ...prev, description: undefined }));
                  }}
                  placeholder="Provide detailed information about the alert..."
                  required
                  rows={6}
                  error={errors.description}
                  description="Use the toolbar to format your message with rich text, links, lists, and more."
                />
              </SharePointSection>

              <SharePointSection title="Alert Configuration">
                <div className={styles.configGrid}>
                  <SharePointSelect
                    label="Alert Type"
                    value={newAlert.AlertType}
                    onChange={(value) => {
                      setNewAlert(prev => ({ ...prev, AlertType: value }));
                      if (errors.AlertType) setErrors(prev => ({ ...prev, AlertType: undefined }));
                    }}
                    options={alertTypeOptions}
                    placeholder="Choose alert style"
                    required
                    error={errors.AlertType}
                    description="Determines the visual appearance of the alert"
                  />

                  <SharePointSelect
                    label="Priority Level"
                    value={newAlert.priority}
                    onChange={(value) => setNewAlert(prev => ({ ...prev, priority: value as AlertPriority }))}
                    options={priorityOptions}
                    description="Higher priorities get more prominent styling"
                  />
                </div>

                <SharePointSelect
                  label="Notifications"
                  value={newAlert.notificationType}
                  onChange={(value) => setNewAlert(prev => ({ ...prev, notificationType: value as NotificationType }))}
                  options={notificationOptions}
                  description="How users should be notified about this alert"
                />

                <SharePointToggle
                  label="Pin Alert to Top"
                  checked={newAlert.isPinned}
                  onChange={(checked) => setNewAlert(prev => ({ ...prev, isPinned: checked }))}
                  description="Pinned alerts stay visible and appear before other alerts"
                />
              </SharePointSection>

              <SharePointSection title="Distribution & Targeting">
                <div className={styles.siteTargeting}>
                  <SiteSelector
                    selectedSites={newAlert.targetSites}
                    onSitesChange={(siteIds) => {
                      setNewAlert(prev => ({ ...prev, targetSites: siteIds }));
                      if (errors.targetSites) setErrors(prev => ({ ...prev, targetSites: undefined }));
                    }}
                    siteDetector={siteDetector}
                    graphClient={graphClient}
                    allowMultiple={true}
                    showPermissionStatus={true}
                  />
                  {errors.targetSites && (
                    <div className={styles.error}>{errors.targetSites}</div>
                  )}
                </div>

                <div className={styles.schedulingOptions}>
                  <h4>Scheduling (Optional)</h4>
                  <div className={styles.dateGrid}>
                    <SharePointInput
                      label="Start Date & Time"
                      value={newAlert.scheduledStart ? newAlert.scheduledStart.toISOString().slice(0, 16) : ""}
                      onChange={(value) => {
                        const date = value ? new Date(value) : undefined;
                        setNewAlert(prev => ({ ...prev, scheduledStart: date }));
                        if (errors.scheduledStart) setErrors(prev => ({ ...prev, scheduledStart: undefined }));
                      }}
                      type="datetime-local"
                      error={errors.scheduledStart}
                      description="When should this alert become visible? Leave empty for immediate."
                    />

                    <SharePointInput
                      label="End Date & Time"
                      value={newAlert.scheduledEnd ? newAlert.scheduledEnd.toISOString().slice(0, 16) : ""}
                      onChange={(value) => {
                        const date = value ? new Date(value) : undefined;
                        setNewAlert(prev => ({ ...prev, scheduledEnd: date }));
                        if (errors.scheduledEnd) setErrors(prev => ({ ...prev, scheduledEnd: undefined }));
                      }}
                      type="datetime-local"
                      error={errors.scheduledEnd}
                      description="When should this alert automatically expire?"
                    />
                  </div>
                </div>
              </SharePointSection>

              {/* Creation Progress */}
              {creationProgress.length > 0 && (
                <SharePointSection title="Creation Status">
                  <div className={styles.creationProgress}>
                    {creationProgress.map((progress, index) => (
                      <div key={index} className={`${styles.progressItem} ${progress.canCreateAlerts ? styles.success : styles.warning}`}>
                        <div className={styles.progressSite}>
                          <strong>{progress.siteName}</strong>
                          <span className={styles.permissionLevel}>({progress.permissionLevel})</span>
                        </div>
                        <div className={styles.progressStatus}>
                          {progress.canCreateAlerts ? '✅ Ready' : '⚠️ Read-only'}
                        </div>
                      </div>
                    ))}
                  </div>
                </SharePointSection>
              )}

              <SharePointSection title="Optional Link" collapsed>
                <div className={styles.configGrid}>
                  <SharePointInput
                    label="Link URL"
                    value={newAlert.linkUrl}
                    onChange={(value) => {
                      setNewAlert(prev => ({ ...prev, linkUrl: value }));
                      if (errors.linkUrl) setErrors(prev => ({ ...prev, linkUrl: undefined }));
                    }}
                    placeholder="https://example.com"
                    type="url"
                    error={errors.linkUrl}
                    description="Optional action link for more information"
                  />

                  <SharePointInput
                    label="Link Text"
                    value={newAlert.linkDescription}
                    onChange={(value) => {
                      setNewAlert(prev => ({ ...prev, linkDescription: value }));
                      if (errors.linkDescription) setErrors(prev => ({ ...prev, linkDescription: undefined }));
                    }}
                    placeholder="Learn more, View details, etc."
                    error={errors.linkDescription}
                    description="Text to display for the link"
                  />
                </div>
              </SharePointSection>

              <div className={styles.formActions}>
                <SharePointButton
                  variant="primary"
                  icon={<Save24Regular />}
                  onClick={handleCreateAlert}
                  disabled={isCreatingAlert}
                >
                  {isCreatingAlert ? 'Creating Alert...' : 'Create Alert'}
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={resetForm}
                >
                  Reset Form
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={() => setShowTemplates(true)}
                >
                  Choose Template
                </SharePointButton>
              </div>
            </div>

            {showPreview && getSelectedAlertType() && (
              <div className={styles.previewColumn}>
                <div className={styles.previewSticky}>
                  <div className={styles.previewToggle}>
                    <SharePointButton
                      variant="secondary"
                      icon={<Eye24Regular />}
                      onClick={() => setShowPreview(!showPreview)}
                    >
                      {showPreview ? "Hide Preview" : "Show Preview"}
                    </SharePointButton>
                  </div>

                  <AlertPreview
                    title={newAlert.title || "Alert Title"}
                    description={newAlert.description || "Alert description will appear here..."}
                    alertType={getSelectedAlertType()!}
                    priority={newAlert.priority}
                    isPinned={newAlert.isPinned}
                    linkUrl={newAlert.linkUrl}
                    linkDescription={newAlert.linkDescription}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );

  const renderManageAlerts = () => (
    <div className={styles.tabContent}>
      <div className={styles.tabHeader}>
        <div>
          <h3>Manage Alerts</h3>
          <p>View, edit, and manage existing alerts across your sites</p>
        </div>
        <div style={{ display: 'flex', gap: '12px' }}>
          {selectedAlerts.length > 0 && (
            <SharePointButton
              variant="danger"
              icon={<Delete24Regular />}
              onClick={handleBulkDelete}
            >
              Delete Selected ({selectedAlerts.length})
            </SharePointButton>
          )}
          <SharePointButton
            variant="primary"
            icon={<Add24Regular />}
            onClick={handleCreateLists}
            disabled={isCreatingLists || isLoadingAlerts}
          >
            {isCreatingLists ? 'Creating Lists...' : 'Create Lists'}
          </SharePointButton>
          <SharePointButton
            variant="secondary"
            onClick={loadExistingAlerts}
            disabled={isLoadingAlerts || isCreatingLists}
          >
            {isLoadingAlerts ? 'Refreshing...' : 'Refresh'}
          </SharePointButton>
        </div>
      </div>

      {isLoadingAlerts ? (
        <div style={{ textAlign: 'center', padding: '40px', color: '#605e5c' }}>
          <div style={{ fontSize: '16px', marginBottom: '8px' }}>Loading alerts...</div>
          <div style={{ fontSize: '14px' }}>Please wait while we fetch your alerts</div>
        </div>
      ) : existingAlerts.length === 0 ? (
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}>📢</div>
          <h4>No Alerts Found</h4>
          <p>No alerts are currently available. This might be because:</p>
          <ul style={{ textAlign: 'left', marginBottom: '20px' }}>
            <li>The Alert Banner lists haven't been created yet</li>
            <li>You don't have access to the lists</li>
            <li>No alerts have been created yet</li>
          </ul>
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
            <SharePointButton
              variant="primary"
              icon={<Add24Regular />}
              onClick={handleCreateLists}
              disabled={isCreatingLists}
            >
              {isCreatingLists ? 'Creating Lists...' : 'Create Alert Lists'}
            </SharePointButton>
            <SharePointButton
              variant="secondary"
              onClick={() => setActiveTab("create")}
            >
              Create First Alert
            </SharePointButton>
          </div>
        </div>
      ) : (
        <div className={styles.alertsList}>
          {existingAlerts.map((alert) => {
            const alertType = alertTypes.find(type => type.name === alert.AlertType);
            const isSelected = selectedAlerts.includes(alert.id);

            return (
              <div key={alert.id} className={`${styles.alertCard} ${isSelected ? styles.selected : ''}`}>
                <div className={styles.alertCardHeader}>
                  <input
                    type="checkbox"
                    checked={isSelected}
                    onChange={(e) => {
                      if (e.target.checked) {
                        setSelectedAlerts(prev => [...prev, alert.id]);
                      } else {
                        setSelectedAlerts(prev => prev.filter(id => id !== alert.id));
                      }
                    }}
                    className={styles.alertCheckbox}
                  />
                  <div className={styles.alertStatus}>
                    <span className={`${styles.statusBadge} ${alert.status.toLowerCase() === 'active' ? styles.active : alert.status.toLowerCase() === 'expired' ? styles.expired : styles.scheduled}`}>
                      {alert.status}
                    </span>
                    {alert.isPinned && (
                      <span className={styles.pinnedBadge}>📌 PINNED</span>
                    )}
                  </div>
                </div>

                <div className={styles.alertCardContent}>
                  {alertType && (
                    <div className={styles.alertTypeIndicator} style={{
                      backgroundColor: alertType.backgroundColor,
                      color: alertType.textColor
                    }}>
                      {alert.AlertType}
                    </div>
                  )}

                  <h4 className={styles.alertCardTitle}>{alert.title}</h4>
                  <div className={styles.alertCardDescription}
                    dangerouslySetInnerHTML={{ __html: alert.description }} />

                  <div className={styles.alertMetadata}>
                    <div className={styles.metadataRow}>
                      <span className={styles.metadataLabel}>Priority:</span>
                      <span className={`${styles.priorityBadge} ${alert.priority.toLowerCase() === 'critical' ? styles.critical :
                          alert.priority.toLowerCase() === 'high' ? styles.high :
                            alert.priority.toLowerCase() === 'medium' ? styles.medium :
                              styles.low
                        }`}>
                        {alert.priority.toUpperCase()}
                      </span>
                    </div>
                    <div className={styles.metadataRow}>
                      <span className={styles.metadataLabel}>Created by:</span>
                      <span>{alert.createdBy}</span>
                    </div>
                    <div className={styles.metadataRow}>
                      <span className={styles.metadataLabel}>Created:</span>
                      <span>{new Date(alert.createdDate).toLocaleDateString()}</span>
                    </div>
                    {alert.scheduledEnd && (
                      <div className={styles.metadataRow}>
                        <span className={styles.metadataLabel}>Expires:</span>
                        <span>{new Date(alert.scheduledEnd).toLocaleDateString()}</span>
                      </div>
                    )}
                  </div>
                </div>

                <div className={styles.alertCardActions}>
                  <SharePointButton
                    variant="secondary"
                    onClick={() => handleEditAlert(alert)}
                  >
                    Edit
                  </SharePointButton>
                  <SharePointButton
                    variant="danger"
                    icon={<Delete24Regular />}
                    onClick={() => handleDeleteAlert(alert.id)}
                  >
                    Delete
                  </SharePointButton>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Edit Alert Dialog */}
      {isEditingAlert && editingAlert && (
        <SharePointDialog
          isOpen={isEditingAlert}
          onClose={handleCancelEdit}
          title={`Edit Alert: ${editingAlert.title}`}
          width={900}
          height={700}
          footer={
            <div className={styles.dialogFooter}>
              <SharePointButton
                variant="primary"
                icon={<Save24Regular />}
                onClick={handleSaveEditedAlert}
              >
                Save Changes
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={handleCancelEdit}
              >
                Cancel
              </SharePointButton>
            </div>
          }
        >
          <div className={styles.editAlertForm}>
            <div className={styles.formWithPreview}>
              <div className={styles.formColumn}>
                <SharePointSection title="Basic Information">
                  <SharePointInput
                    label="Alert Title"
                    value={editingAlert.title}
                    onChange={(value) => setEditingAlert(prev => prev ? { ...prev, title: value } : null)}
                    placeholder="Enter a clear, concise title"
                    required
                  />

                  <SharePointRichTextEditor
                    label="Alert Description"
                    value={editingAlert.description}
                    onChange={(value) => setEditingAlert(prev => prev ? { ...prev, description: value } : null)}
                    placeholder="Provide detailed information about the alert..."
                    required
                    rows={6}
                  />
                </SharePointSection>

                <SharePointSection title="Alert Configuration">
                  <div className={styles.configGrid}>
                    <SharePointSelect
                      label="Alert Type"
                      value={editingAlert.AlertType}
                      onChange={(value) => setEditingAlert(prev => prev ? { ...prev, AlertType: value } : null)}
                      options={alertTypeOptions}
                      placeholder="Choose alert style"
                      required
                    />

                    <SharePointSelect
                      label="Priority Level"
                      value={editingAlert.priority}
                      onChange={(value) => setEditingAlert(prev => prev ? { ...prev, priority: value as AlertPriority } : null)}
                      options={priorityOptions}
                    />
                  </div>

                  <SharePointSelect
                    label="Notifications"
                    value={editingAlert.notificationType || NotificationType.None}
                    onChange={(value) => setEditingAlert(prev => prev ? { ...prev, notificationType: value as NotificationType } : null)}
                    options={notificationOptions}
                  />

                  <SharePointToggle
                    label="Pin Alert to Top"
                    checked={editingAlert.isPinned}
                    onChange={(checked) => setEditingAlert(prev => prev ? { ...prev, isPinned: checked } : null)}
                  />
                </SharePointSection>

                <SharePointSection title="Scheduling">
                  <div className={styles.dateGrid}>
                    <SharePointInput
                      label="Start Date & Time"
                      value={editingAlert.scheduledStart ? editingAlert.scheduledStart.toISOString().slice(0, 16) : ""}
                      onChange={(value) => {
                        const date = value ? new Date(value) : undefined;
                        setEditingAlert(prev => prev ? { ...prev, scheduledStart: date } : null);
                      }}
                      type="datetime-local"
                    />

                    <SharePointInput
                      label="End Date & Time"
                      value={editingAlert.scheduledEnd ? editingAlert.scheduledEnd.toISOString().slice(0, 16) : ""}
                      onChange={(value) => {
                        const date = value ? new Date(value) : undefined;
                        setEditingAlert(prev => prev ? { ...prev, scheduledEnd: date } : null);
                      }}
                      type="datetime-local"
                    />
                  </div>
                </SharePointSection>
              </div>

              {(() => {
                const selectedAlertType = alertTypes.find(type => type.name === editingAlert.AlertType);
                return selectedAlertType ? (
                  <div className={styles.previewColumn}>
                    <div className={styles.previewSticky}>
                      <AlertPreview
                        title={editingAlert.title || "Alert Title"}
                        description={editingAlert.description || "Alert description will appear here..."}
                        alertType={selectedAlertType}
                        priority={editingAlert.priority}
                        isPinned={editingAlert.isPinned}
                        linkUrl={editingAlert.linkUrl}
                        linkDescription={editingAlert.linkDescription}
                      />
                    </div>
                  </div>
                ) : null;
              })()}
            </div>
          </div>
        </SharePointDialog>
      )}
    </div>
  );

  const renderAlertTypes = () => (
    <div className={styles.tabContent}>
      <div className={styles.tabHeader}>
        <div>
          <h3>Manage Alert Types</h3>
          <p>Create and customize the visual appearance of different alert categories</p>
        </div>
        <SharePointButton
          variant="primary"
          icon={<Add24Regular />}
          onClick={() => setIsCreatingType(true)}
        >
          Create New Type
        </SharePointButton>
      </div>

      {isCreatingType && (
        <SharePointSection title="Create New Alert Type">
          <div className={styles.typeFormWithPreview}>
            <div className={styles.typeFormColumn}>
              <SharePointInput
                label="Type Name"
                value={newAlertType.name}
                onChange={(value) => setNewAlertType(prev => ({ ...prev, name: value }))}
                placeholder="e.g., Maintenance, Emergency, Update"
                required
                description="A unique name for this alert type"
              />

              <div className={styles.colorRow}>
                <ColorPicker
                  label="Background Color"
                  value={newAlertType.backgroundColor}
                  onChange={(color) => setNewAlertType(prev => ({ ...prev, backgroundColor: color }))}
                  description="Main background color for alerts of this type"
                />

                <ColorPicker
                  label="Text Color"
                  value={newAlertType.textColor}
                  onChange={(color) => setNewAlertType(prev => ({ ...prev, textColor: color }))}
                  description="Text color that contrasts well with background"
                />
              </div>

              <SharePointInput
                label="Icon Name"
                value={newAlertType.iconName}
                onChange={(value) => setNewAlertType(prev => ({ ...prev, iconName: value }))}
                placeholder="Info, Warning, Error, CheckmarkCircle, etc."
                description="Fluent UI icon name (optional)"
              />

              <SharePointTextArea
                label="Custom CSS Styles"
                value={newAlertType.additionalStyles || ""}
                onChange={(value) => setNewAlertType(prev => ({ ...prev, additionalStyles: value }))}
                placeholder="Additional CSS styles (advanced)"
                rows={3}
                description="Optional custom CSS for advanced styling"
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
            </div>

            <div className={styles.typePreviewColumn}>
              <AlertPreview
                title="Sample Alert Title"
                description="This is how alerts with this type will appear to users. The preview updates as you change the colors and settings."
                alertType={newAlertType}
                priority={AlertPriority.Medium}
                isPinned={false}
              />
            </div>
          </div>
        </SharePointSection>
      )}

      <SharePointSection title="Existing Alert Types">
        <div className={styles.dragDropInstructions}>
          <p>💡 <strong>Tip:</strong> Drag and drop alert types to reorder them. The order here determines the display order in dropdown menus.</p>
        </div>
        <div className={styles.existingTypes}>
          {alertTypes.map((type, index) => (
            <div
              key={type.name}
              className={styles.alertTypeCard}
              draggable
              onDragStart={(e) => handleDragStart(e, index)}
              onDragEnd={handleDragEnd}
              onDragOver={handleDragOver}
              onDrop={(e) => handleDrop(e, index)}
            >
              <div className={styles.dragHandle}>
                <span className={styles.dragIcon}>⋮⋮</span>
                <span className={styles.orderNumber}>#{index + 1}</span>
              </div>

              <div className={styles.typePreview}>
                <AlertPreview
                  title={`Sample ${type.name} Alert`}
                  description="This is a preview of how this alert type appears."
                  alertType={type}
                  priority={AlertPriority.Medium}
                  isPinned={false}
                />
              </div>

              <div className={styles.typeActions}>
                <SharePointButton
                  variant="danger"
                  icon={<Delete24Regular />}
                  onClick={() => handleDeleteAlertType(index)}
                >
                  Delete
                </SharePointButton>
              </div>
            </div>
          ))}

          {alertTypes.length === 0 && (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>🎨</div>
              <h4>No Alert Types</h4>
              <p>Create your first alert type to get started with customized alert styling.</p>
            </div>
          )}
        </div>
      </SharePointSection>
    </div>
  );

  const renderSettings = () => (
    <div className={styles.tabContent}>
      <SharePointSection title="Feature Settings">
        <div className={styles.settingsGrid}>
          <SharePointToggle
            label="Enable User Targeting"
            checked={settings.userTargetingEnabled}
            onChange={(checked) => setSettings(prev => ({ ...prev, userTargetingEnabled: checked }))}
            description="Allow alerts to target specific users or groups based on SharePoint profiles and security groups"
          />

          <SharePointToggle
            label="Enable Browser Notifications"
            checked={settings.notificationsEnabled}
            onChange={(checked) => setSettings(prev => ({ ...prev, notificationsEnabled: checked }))}
            description="Send native browser notifications for critical and high-priority alerts to ensure visibility"
          />

          <SharePointToggle
            label="Enable Rich Media Support"
            checked={settings.richMediaEnabled}
            onChange={(checked) => setSettings(prev => ({ ...prev, richMediaEnabled: checked }))}
            description="Support images, videos, HTML content, and markdown formatting in alert descriptions"
          />
        </div>
      </SharePointSection>

      <SharePointSection title="SharePoint Setup">
        <div className={styles.settingsGrid}>
          <div style={{ gridColumn: '1 / -1' }}>
            <p style={{ marginBottom: '16px', color: '#605e5c' }}>
              The Alert Banner system requires two SharePoint lists: <strong>Alerts</strong> (for storing alert content) and <strong>AlertBannerTypes</strong> (for alert styling configurations).
            </p>
            
            <div style={{ display: 'flex', gap: '12px', alignItems: 'center', flexWrap: 'wrap' }}>
              <SharePointButton
                variant="primary"
                icon={<Add24Regular />}
                onClick={handleCreateLists}
                disabled={isCreatingLists}
              >
                {isCreatingLists ? 'Creating Lists...' : 'Create/Initialize Lists'}
              </SharePointButton>
              
              <div style={{ fontSize: '14px', color: '#605e5c' }}>
                Creates the required SharePoint lists on the current site if they don't exist.
              </div>
            </div>
          </div>
        </div>
      </SharePointSection>

      <LanguageFieldManager
        alertService={alertService}
        onLanguageChange={(languages) => {
          console.log('Active languages updated:', languages);
          // Optionally store active languages in settings
        }}
      />

      <SharePointSection title="System Information" collapsed>
        <div className={styles.systemInfo}>
          <div className={styles.infoGrid}>
            <div className={styles.infoItem}>
              <span className={styles.infoLabel}>Total Alert Types:</span>
              <span className={styles.infoValue}>{alertTypes.length}</span>
            </div>
            <div className={styles.infoItem}>
              <span className={styles.infoLabel}>User Targeting:</span>
              <span className={`${styles.infoValue} ${settings.userTargetingEnabled ? styles.enabled : styles.disabled}`}>
                {settings.userTargetingEnabled ? "Enabled" : "Disabled"}
              </span>
            </div>
            <div className={styles.infoItem}>
              <span className={styles.infoLabel}>Notifications:</span>
              <span className={`${styles.infoValue} ${settings.notificationsEnabled ? styles.enabled : styles.disabled}`}>
                {settings.notificationsEnabled ? "Enabled" : "Disabled"}
              </span>
            </div>
            <div className={styles.infoItem}>
              <span className={styles.infoLabel}>Rich Media:</span>
              <span className={`${styles.infoValue} ${settings.richMediaEnabled ? styles.enabled : styles.disabled}`}>
                {settings.richMediaEnabled ? "Enabled" : "Disabled"}
              </span>
            </div>
          </div>
        </div>
      </SharePointSection>
    </div>
  );

  const footer = (
    <div className={styles.dialogFooter}>
      <SharePointButton variant="primary" onClick={handleSaveSettings}>
        Save All Settings
      </SharePointButton>
      <SharePointButton variant="secondary" onClick={() => setIsOpen(false)}>
        Close
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
        title="Alert Banner Configuration"
        width={1100}
        height={800}
        footer={footer}
      >
        <div className={styles.settingsContainer}>
          <div className={styles.tabs}>
            <button
              className={`${styles.tab} ${activeTab === "create" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("create")}
            >
              <span className={styles.tabIcon}>➕</span>
              Create Alert
            </button>
            <button
              className={`${styles.tab} ${activeTab === "manage" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("manage")}
            >
              <span className={styles.tabIcon}>📋</span>
              Manage Alerts
            </button>
            <button
              className={`${styles.tab} ${activeTab === "types" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("types")}
            >
              <span className={styles.tabIcon}>🎨</span>
              Alert Types
            </button>
            <button
              className={`${styles.tab} ${activeTab === "settings" ? styles.activeTab : ""}`}
              onClick={() => setActiveTab("settings")}
            >
              <span className={styles.tabIcon}>⚙️</span>
              Settings
            </button>
          </div>

          {activeTab === "create" && renderCreateAlert()}
          {activeTab === "manage" && renderManageAlerts()}
          {activeTab === "types" && renderAlertTypes()}
          {activeTab === "settings" && renderSettings()}
        </div>
      </SharePointDialog>
    </>
  );
};

export default AlertSettings;