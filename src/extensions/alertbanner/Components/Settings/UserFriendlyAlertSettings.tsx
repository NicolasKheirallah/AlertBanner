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
import styles from "./UserFriendlyAlertSettings.module.scss";

export interface IUserFriendlyAlertSettingsProps {
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
  includeSubsites: boolean;
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

const UserFriendlyAlertSettings: React.FC<IUserFriendlyAlertSettingsProps> = ({
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
  const [activeTab, setActiveTab] = React.useState<"create" | "types" | "settings">("create");
  const [showTemplates, setShowTemplates] = React.useState(true);
  const [isCreatingType, setIsCreatingType] = React.useState(false);
  const [showPreview, setShowPreview] = React.useState(true);
  
  // Site context and targeting
  const [siteDetector] = React.useState(() => new SiteContextDetector(graphClient, context));
  const [currentSiteContext, setCurrentSiteContext] = React.useState<ISiteContext | null>(null);
  const [creationProgress, setCreationProgress] = React.useState<ISiteValidationResult[]>([]);
  const [isCreatingAlert, setIsCreatingAlert] = React.useState(false);
  
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
    linkDescription: "",
    targetSites: [],
    includeSubsites: false,
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
      includeSubsites: prev.includeSubsites,
      scheduledStart: prev.scheduledStart,
      scheduledEnd: prev.scheduledEnd
    }));
    setShowTemplates(false);
    setErrors({});
  };

  const handleSaveSettings = () => {
    const updatedSettings = {
      ...settings,
      alertTypesJson: JSON.stringify(alertTypes, null, 2)
    };
    onSettingsChange(updatedSettings);
    setIsOpen(false);
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

      // Create alert on each valid site
      const createPromises = validSites.map(async (siteValidation) => {
        try {
          const alertItem = {
            title: newAlert.title.trim(),
            description: newAlert.description.trim(),
            AlertType: newAlert.AlertType,
            priority: newAlert.priority,
            isPinned: newAlert.isPinned,
            notificationType: newAlert.notificationType,
            ...(newAlert.linkUrl && newAlert.linkDescription && {
              link: {
                Url: newAlert.linkUrl.trim(),
                Description: newAlert.linkDescription.trim()
              }
            }),
            createdDate: newAlert.scheduledStart?.toISOString() || new Date().toISOString(),
            createdBy: context.pageContext.user.displayName || "Alert Settings",
            // Additional metadata for multi-site deployment
            metadata: {
              sourceSiteId: currentSiteContext?.siteId,
              sourceSiteName: currentSiteContext?.siteName,
              endDate: newAlert.scheduledEnd?.toISOString(),
              deploymentTargets: newAlert.targetSites
            }
          };

          // Here you would actually create the alert in SharePoint
          // For now, we'll simulate the creation
          console.log('Creating alert item:', alertItem);
          await new Promise(resolve => setTimeout(resolve, 1000)); // Simulate API call
          
          return {
            siteId: siteValidation.siteId,
            success: true,
            siteName: siteValidation.siteName
          };
        } catch (error) {
          return {
            siteId: siteValidation.siteId,
            success: false,
            siteName: siteValidation.siteName,
            error: error.message
          };
        }
      });

      const results = await Promise.allSettled(createPromises);
      const finalResults = results.map(result => 
        result.status === 'fulfilled' ? result.value : {
          siteId: '',
          success: false,
          siteName: 'Unknown',
          error: 'Creation failed'
        }
      );

      const successCount = finalResults.filter(r => r.success).length;
      
      // Show comprehensive results
      const alertElement = document.createElement('div');
      alertElement.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${successCount > 0 ? '#107c10' : '#d13438'};
        color: white;
        padding: 16px 20px;
        border-radius: 4px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        font-family: Segoe UI;
        font-size: 14px;
        max-width: 300px;
      `;
      
      if (successCount === finalResults.length) {
        alertElement.textContent = `✅ Alert created successfully on ${successCount} site${successCount !== 1 ? 's' : ''}!`;
      } else if (successCount > 0) {
        alertElement.textContent = `⚠️ Alert created on ${successCount} of ${finalResults.length} sites. Check details for failed sites.`;
      } else {
        alertElement.textContent = `❌ Failed to create alert on any sites. Please check your permissions.`;
      }
      
      document.body.appendChild(alertElement);
      
      setTimeout(() => {
        if (document.body.contains(alertElement)) {
          document.body.removeChild(alertElement);
        }
      }, 5000);
      
      // Reset form if all succeeded
      if (successCount === finalResults.length) {
        resetForm();
      }
      
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
      alertElement.textContent = `❌ ${error.message}`;
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
      includeSubsites: false,
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
        <div className={styles.existingTypes}>
          {alertTypes.map((type, index) => (
            <div key={type.name} className={styles.alertTypeCard}>
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
          {activeTab === "types" && renderAlertTypes()}
          {activeTab === "settings" && renderSettings()}
        </div>
      </SharePointDialog>
    </>
  );
};

export default UserFriendlyAlertSettings;