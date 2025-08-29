import * as React from "react";
import { Save24Regular, Eye24Regular, Dismiss24Regular } from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
  ISharePointSelectOption
} from "../../UI/SharePointControls";
import SharePointRichTextEditor from "../../UI/SharePointRichTextEditor";
import AlertPreview from "../../UI/AlertPreview";
import AlertTemplates, { IAlertTemplate } from "../../UI/AlertTemplates";
import SiteSelector from "../../UI/SiteSelector";
import { AlertPriority, NotificationType, IAlertType } from "../../Alerts/IAlerts";
import { SiteContextDetector, ISiteValidationResult } from "../../Utils/SiteContextDetector";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "../AlertSettings.module.scss";

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
}

export interface IFormErrors {
  title?: string;
  description?: string;
  AlertType?: string;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
}

export interface ICreateAlertTabProps {
  newAlert: INewAlert;
  setNewAlert: React.Dispatch<React.SetStateAction<INewAlert>>;
  errors: IFormErrors;
  setErrors: React.Dispatch<React.SetStateAction<IFormErrors>>;
  alertTypes: IAlertType[];
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
  siteDetector: SiteContextDetector;
  alertService: SharePointAlertService;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  creationProgress: ISiteValidationResult[];
  setCreationProgress: React.Dispatch<React.SetStateAction<ISiteValidationResult[]>>;
  isCreatingAlert: boolean;
  setIsCreatingAlert: React.Dispatch<React.SetStateAction<boolean>>;
  showPreview: boolean;
  setShowPreview: React.Dispatch<React.SetStateAction<boolean>>;
  showTemplates: boolean;
  setShowTemplates: React.Dispatch<React.SetStateAction<boolean>>;
}

const CreateAlertTab: React.FC<ICreateAlertTabProps> = ({
  newAlert,
  setNewAlert,
  errors,
  setErrors,
  alertTypes,
  userTargetingEnabled,
  notificationsEnabled,
  richMediaEnabled,
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
  setShowTemplates
}) => {
  // Priority options
  const priorityOptions: ISharePointSelectOption[] = [
    { value: AlertPriority.Low, label: "Low Priority - Informational updates" },
    { value: AlertPriority.Medium, label: "Medium Priority - General announcements" },
    { value: AlertPriority.High, label: "High Priority - Important updates" },
    { value: AlertPriority.Critical, label: "Critical Priority - Urgent action required" }
  ];

  // Notification type options
  const notificationOptions: ISharePointSelectOption[] = [
    { value: NotificationType.None, label: "None" },
    { value: NotificationType.Browser, label: "Browser Notification" },
    { value: NotificationType.Email, label: "Email Notification" },
    { value: NotificationType.Both, label: "Browser + Email" }
  ];

  // Alert type options
  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(type => ({
    value: type.name,
    label: type.name
  }));

  const handleTemplateSelect = React.useCallback((template: IAlertTemplate) => {
    setNewAlert(prev => ({
      ...prev,
      title: template.template.title,
      description: template.template.description,
      priority: template.template.priority,
      notificationType: template.template.notificationType,
      isPinned: template.template.isPinned,
      linkUrl: template.template.linkUrl || "",
      linkDescription: template.template.linkDescription || ""
    }));
    setShowTemplates(false);
  }, [setNewAlert, setShowTemplates]);

  const validateForm = React.useCallback((): boolean => {
    const newErrors: IFormErrors = {};

    if (!newAlert.title?.trim()) {
      newErrors.title = "Title is required";
    } else if (newAlert.title.length < 3) {
      newErrors.title = "Title must be at least 3 characters";
    } else if (newAlert.title.length > 100) {
      newErrors.title = "Title cannot exceed 100 characters";
    }

    if (!newAlert.description?.trim()) {
      newErrors.description = "Description is required";
    } else if (newAlert.description.length < 10) {
      newErrors.description = "Description must be at least 10 characters";
    }

    if (!newAlert.AlertType) {
      newErrors.AlertType = "Alert type is required";
    }

    if (newAlert.linkUrl && newAlert.linkUrl.trim()) {
      try {
        new URL(newAlert.linkUrl);
      } catch {
        newErrors.linkUrl = "Please enter a valid URL";
      }
    }

    if (newAlert.linkUrl && !newAlert.linkDescription?.trim()) {
      newErrors.linkDescription = "Link description is required when URL is provided";
    }

    if (newAlert.targetSites.length === 0) {
      newErrors.targetSites = "At least one target site must be selected";
    }

    if (newAlert.scheduledStart && newAlert.scheduledEnd) {
      if (newAlert.scheduledStart >= newAlert.scheduledEnd) {
        newErrors.scheduledEnd = "End date must be after start date";
      }
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }, [newAlert, setErrors]);

  const handleCreateAlert = React.useCallback(async () => {
    if (!validateForm()) return;

    setIsCreatingAlert(true);
    setCreationProgress([]);

    try {
      // Transform INewAlert to match service expectations
      const alertData = {
        title: newAlert.title,
        description: newAlert.description,
        AlertType: newAlert.AlertType,
        priority: newAlert.priority,
        isPinned: newAlert.isPinned,
        notificationType: newAlert.notificationType,
        linkUrl: newAlert.linkUrl,
        linkDescription: newAlert.linkDescription,
        targetSites: newAlert.targetSites,
        scheduledStart: newAlert.scheduledStart?.toISOString(),
        scheduledEnd: newAlert.scheduledEnd?.toISOString()
      };

      await alertService.createAlert(alertData);
      
      // Success - create a simple success result
      setCreationProgress([{
        siteId: "success",
        siteName: "Alert Created",
        hasAccess: true,
        canCreateAlerts: true,
        permissionLevel: "success",
        error: ""
      }]);

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
        scheduledEnd: undefined
      });
      setShowTemplates(true);
    } catch (error) {
      console.error('Error creating alert:', error);
      setCreationProgress([{
        siteId: "error",
        siteName: "Creation Error",
        hasAccess: false,
        canCreateAlerts: false,
        permissionLevel: "error",
        error: error instanceof Error ? error.message : "Unknown error occurred"
      }]);
    } finally {
      setIsCreatingAlert(false);
    }
  }, [validateForm, setIsCreatingAlert, setCreationProgress, alertService, newAlert, setNewAlert, alertTypes, setShowTemplates]);

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
      scheduledEnd: undefined
    });
    setErrors({});
    setShowTemplates(true);
  }, [setNewAlert, alertTypes, setErrors, setShowTemplates]);

  const getCurrentAlertType = React.useCallback((): IAlertType | undefined => {
    return alertTypes.find(type => type.name === newAlert.AlertType);
  }, [alertTypes, newAlert.AlertType]);

  return (
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
                  error={errors.description}
                  description="Use the toolbar to format your message with rich text, links, lists, and more."
                />
              </SharePointSection>

              <SharePointSection title="Alert Configuration">
                <SharePointSelect
                  label="Alert Type"
                  value={newAlert.AlertType}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, AlertType: value }));
                    if (errors.AlertType) setErrors(prev => ({ ...prev, AlertType: undefined }));
                  }}
                  options={alertTypeOptions}
                  required
                  error={errors.AlertType}
                  description="Choose the visual style and importance level"
                />

                <SharePointSelect
                  label="Priority Level"
                  value={newAlert.priority}
                  onChange={(value) => setNewAlert(prev => ({ ...prev, priority: value as AlertPriority }))}
                  options={priorityOptions}
                  required
                  description="This affects the visual styling and user attention level"
                />

                <SharePointToggle
                  label="Pin Alert"
                  checked={newAlert.isPinned}
                  onChange={(checked) => setNewAlert(prev => ({ ...prev, isPinned: checked }))}
                  description="Pinned alerts stay at the top and are harder to dismiss"
                />

                {notificationsEnabled && (
                  <SharePointSelect
                    label="Notification Type"
                    value={newAlert.notificationType}
                    onChange={(value) => setNewAlert(prev => ({ ...prev, notificationType: value as NotificationType }))}
                    options={notificationOptions}
                    description="How users will be notified about this alert"
                  />
                )}
              </SharePointSection>

              <SharePointSection title="Action Link (Optional)">
                <SharePointInput
                  label="Link URL"
                  value={newAlert.linkUrl}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, linkUrl: value }));
                    if (errors.linkUrl) setErrors(prev => ({ ...prev, linkUrl: undefined }));
                  }}
                  placeholder="https://example.com/more-info"
                  error={errors.linkUrl}
                  description="Optional link for users to get more information or take action"
                />

                {newAlert.linkUrl && (
                  <SharePointInput
                    label="Link Description"
                    value={newAlert.linkDescription}
                    onChange={(value) => {
                      setNewAlert(prev => ({ ...prev, linkDescription: value }));
                      if (errors.linkDescription) setErrors(prev => ({ ...prev, linkDescription: undefined }));
                    }}
                    placeholder="Learn More"
                    required={!!newAlert.linkUrl}
                    error={errors.linkDescription}
                    description="Text that will appear on the action button"
                  />
                )}
              </SharePointSection>

              <SharePointSection title="Target Sites">
                <SiteSelector
                  selectedSites={newAlert.targetSites}
                  onSitesChange={(sites) => {
                    setNewAlert(prev => ({ ...prev, targetSites: sites }));
                    if (errors.targetSites) setErrors(prev => ({ ...prev, targetSites: undefined }));
                  }}
                  siteDetector={siteDetector}
                  graphClient={graphClient}
                  showPermissionStatus={true}
                />
                {errors.targetSites && (
                  <div className={styles.errorMessage}>{errors.targetSites}</div>
                )}
              </SharePointSection>

              <SharePointSection title="Scheduling (Optional)">
                <SharePointInput
                  label="Start Date & Time"
                  type="datetime-local"
                  value={newAlert.scheduledStart ? new Date(newAlert.scheduledStart.getTime() - newAlert.scheduledStart.getTimezoneOffset() * 60000).toISOString().slice(0, 16) : ""}
                  onChange={(value) => {
                    setNewAlert(prev => ({ 
                      ...prev, 
                      scheduledStart: value ? new Date(value) : undefined 
                    }));
                    if (errors.scheduledStart) setErrors(prev => ({ ...prev, scheduledStart: undefined }));
                  }}
                  error={errors.scheduledStart}
                  description="When should this alert become visible? Leave empty to show immediately."
                />

                <SharePointInput
                  label="End Date & Time"
                  type="datetime-local"
                  value={newAlert.scheduledEnd ? new Date(newAlert.scheduledEnd.getTime() - newAlert.scheduledEnd.getTimezoneOffset() * 60000).toISOString().slice(0, 16) : ""}
                  onChange={(value) => {
                    setNewAlert(prev => ({ 
                      ...prev, 
                      scheduledEnd: value ? new Date(value) : undefined 
                    }));
                    if (errors.scheduledEnd) setErrors(prev => ({ ...prev, scheduledEnd: undefined }));
                  }}
                  error={errors.scheduledEnd}
                  description="When should this alert automatically hide? Leave empty to keep it visible until manually removed."
                />
              </SharePointSection>

              <div className={styles.formActions}>
                <SharePointButton
                  variant="primary"
                  onClick={handleCreateAlert}
                  disabled={isCreatingAlert || alertTypes.length === 0}
                  icon={<Save24Regular />}
                >
                  {isCreatingAlert ? "Creating Alert..." : "Create Alert"}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={resetForm}
                  disabled={isCreatingAlert}
                  icon={<Dismiss24Regular />}
                >
                  Reset Form
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={() => setShowPreview(!showPreview)}
                  icon={<Eye24Regular />}
                >
                  {showPreview ? "Hide Preview" : "Show Preview"}
                </SharePointButton>
              </div>

              {/* Creation Progress */}
              {creationProgress.length > 0 && (
                <div className={styles.alertsList}>
                  <h3>Creation Results:</h3>
                  {creationProgress.map((result, index) => (
                    <div
                      key={index}
                      className={`${styles.alertCard} ${result.error ? styles.error : styles.success}`}
                    >
                      <strong>{result.siteName}</strong>: {result.error ? `❌ ${result.error}` : "✅ Created successfully"}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Preview Column */}
            {showPreview && (
              <div className={styles.formColumn}>
                <div className={styles.alertCard}>
                  <h3>Preview</h3>
                  <AlertPreview
                    title={newAlert.title || "Alert Title"}
                    description={newAlert.description || "Alert description will appear here..."}
                    alertType={getCurrentAlertType() || { name: "Default", iconName: "Info", backgroundColor: "#0078d4", textColor: "#ffffff", additionalStyles: "", priorityStyles: {} }}
                    priority={newAlert.priority}
                    linkUrl={newAlert.linkUrl}
                    linkDescription={newAlert.linkDescription}
                    isPinned={newAlert.isPinned}
                  />
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default CreateAlertTab;