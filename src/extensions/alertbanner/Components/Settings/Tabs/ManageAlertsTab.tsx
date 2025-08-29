import * as React from "react";
import { Delete24Regular, Edit24Regular } from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
  ISharePointSelectOption
} from "../../UI/SharePointControls";
import SharePointRichTextEditor from "../../UI/SharePointRichTextEditor";
import SharePointDialog from "../../UI/SharePointDialog";
import { AlertPriority, NotificationType, IAlertType } from "../../Alerts/IAlerts";
import { SiteContextDetector } from "../../Utils/SiteContextDetector";
import { SharePointAlertService, IAlertItem } from "../../Services/SharePointAlertService";
import { htmlSanitizer } from "../../Utils/HtmlSanitizer";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from "../AlertSettings.module.scss";

export interface IEditingAlert extends Omit<IAlertItem, 'scheduledStart' | 'scheduledEnd'> {
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
  setActiveTab: React.Dispatch<React.SetStateAction<"create" | "manage" | "types" | "settings">>;
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
  setActiveTab
}) => {
  const [editErrors, setEditErrors] = React.useState<IFormErrors>({});

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

  const loadExistingAlerts = React.useCallback(async () => {
    setIsLoadingAlerts(true);
    try {
      const alerts = await alertService.getAlerts();
      setExistingAlerts(alerts);
    } catch (error) {
      console.error('Error loading alerts:', error);
      setExistingAlerts([]);
    } finally {
      setIsLoadingAlerts(false);
    }
  }, [alertService]);

  const handleBulkDelete = React.useCallback(async () => {
    if (selectedAlerts.length === 0) return;
    
    if (!confirm(`Are you sure you want to delete ${selectedAlerts.length} alert(s)? This action cannot be undone.`)) {
      return;
    }

    try {
      await Promise.all(
        selectedAlerts.map(alertId => alertService.deleteAlert(alertId))
      );
      
      // Refresh the alerts list
      await loadExistingAlerts();
      setSelectedAlerts([]);
    } catch (error) {
      console.error('Error deleting alerts:', error);
      alert('Failed to delete some alerts. Please try again.');
    }
  }, [selectedAlerts, alertService, loadExistingAlerts]);

  const handleEditAlert = React.useCallback((alert: IAlertItem) => {
    const editingData: IEditingAlert = {
      ...alert,
      scheduledStart: alert.scheduledStart ? new Date(alert.scheduledStart) : undefined,
      scheduledEnd: alert.scheduledEnd ? new Date(alert.scheduledEnd) : undefined
    };
    setEditingAlert(editingData);
    setEditErrors({});
  }, [setEditingAlert]);

  const handleDeleteAlert = React.useCallback(async (alertId: string, alertTitle: string) => {
    if (!confirm(`Are you sure you want to delete "${alertTitle}"? This action cannot be undone.`)) {
      return;
    }

    try {
      await alertService.deleteAlert(alertId);
      await loadExistingAlerts();
    } catch (error) {
      console.error('Error deleting alert:', error);
      alert('Failed to delete alert. Please try again.');
    }
  }, [alertService, loadExistingAlerts]);

  const validateEditForm = React.useCallback((): boolean => {
    if (!editingAlert) return false;

    const newErrors: IFormErrors = {};

    if (!editingAlert.title?.trim()) {
      newErrors.title = "Title is required";
    } else if (editingAlert.title.length < 3) {
      newErrors.title = "Title must be at least 3 characters";
    } else if (editingAlert.title.length > 100) {
      newErrors.title = "Title cannot exceed 100 characters";
    }

    if (!editingAlert.description?.trim()) {
      newErrors.description = "Description is required";
    } else if (editingAlert.description.length < 10) {
      newErrors.description = "Description must be at least 10 characters";
    }

    if (!editingAlert.AlertType) {
      newErrors.AlertType = "Alert type is required";
    }

    if (editingAlert.linkUrl && editingAlert.linkUrl.trim()) {
      try {
        new URL(editingAlert.linkUrl);
      } catch {
        newErrors.linkUrl = "Please enter a valid URL";
      }
    }

    if (editingAlert.linkUrl && !editingAlert.linkDescription?.trim()) {
      newErrors.linkDescription = "Link description is required when URL is provided";
    }

    if (editingAlert.scheduledStart && editingAlert.scheduledEnd) {
      if (editingAlert.scheduledStart >= editingAlert.scheduledEnd) {
        newErrors.scheduledEnd = "End date must be after start date";
      }
    }

    setEditErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }, [editingAlert]);

  const handleSaveEdit = React.useCallback(async () => {
    if (!editingAlert || !validateEditForm()) return;

    setIsEditingAlert(true);
    try {
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
        scheduledEnd: editingAlert.scheduledEnd?.toISOString()
      });

      setEditingAlert(null);
      setEditErrors({});
      await loadExistingAlerts();
    } catch (error) {
      console.error('Error updating alert:', error);
      alert('Failed to update alert. Please try again.');
    } finally {
      setIsEditingAlert(false);
    }
  }, [editingAlert, validateEditForm, setIsEditingAlert, alertService, setEditingAlert, setEditErrors, loadExistingAlerts]);

  const handleCancelEdit = React.useCallback(() => {
    setEditingAlert(null);
    setEditErrors({});
  }, [setEditingAlert]);

  // Load alerts on mount
  React.useEffect(() => {
    loadExistingAlerts();
  }, [loadExistingAlerts]);

  return (
    <>
      <div className={styles.tabContent}>
        <div className={styles.tabHeader}>
          <div>
            <h3>Manage Alerts</h3>
            <p>View, edit, and manage existing alerts across your sites</p>
          </div>
          <div className={styles.flexRowGap12}>
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
              variant="secondary"
              onClick={loadExistingAlerts}
              disabled={isLoadingAlerts}
            >
              {isLoadingAlerts ? 'Refreshing...' : 'Refresh'}
            </SharePointButton>
          </div>
        </div>

        {isLoadingAlerts ? (
          <div className={styles.loadingContainer}>
            <div className={styles.loadingTitle}>Loading alerts...</div>
            <div className={styles.loadingSubtitle}>Please wait while we fetch your alerts</div>
          </div>
        ) : existingAlerts.length === 0 ? (
          <div className={styles.emptyState}>
            <div className={styles.emptyIcon}>📢</div>
            <h4>No Alerts Found</h4>
            <p>No alerts are currently available. This might be because:</p>
            <ul className={styles.emptyStateList}>
              <li>The Alert Banner lists haven't been created yet</li>
              <li>You don't have access to the lists</li>
              <li>No alerts have been created yet</li>
            </ul>
            <div className={styles.flexRowCentered}>
              <SharePointButton
                variant="primary"
                onClick={() => setActiveTab("create")}
              >
                Create First Alert
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={loadExistingAlerts}
              >
                Refresh
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
                      <div 
                        className={styles.alertTypeIndicator}
                        style={{
                          '--bg-color': alertType.backgroundColor,
                          '--text-color': alertType.textColor
                        } as React.CSSProperties}
                      >
                        {alert.AlertType}
                      </div>
                    )}
                    <h4 className={styles.alertCardTitle}>{alert.title}</h4>
                    <div className={styles.alertCardDescription}
                      dangerouslySetInnerHTML={{ 
                        __html: htmlSanitizer.sanitizeHtml(alert.description?.substring(0, 150) + (alert.description?.length > 150 ? '...' : ''))
                      }}
                    />
                    
                    <div className={styles.alertMetaData}>
                      <div className={styles.metaInfo}>
                        <strong>Priority:</strong> {alert.priority}
                      </div>
                      {alert.linkUrl && (
                        <div className={styles.metaInfo}>
                          <strong>Action:</strong> {alert.linkDescription}
                        </div>
                      )}
                      <div className={styles.metaInfo}>
                        <strong>Created:</strong> {new Date(alert.createdDate || Date.now()).toLocaleDateString()}
                      </div>
                      {alert.scheduledStart && (
                        <div className={styles.metaInfo}>
                          <strong>Start:</strong> {new Date(alert.scheduledStart).toLocaleString()}
                        </div>
                      )}
                      {alert.scheduledEnd && (
                        <div className={styles.metaInfo}>
                          <strong>End:</strong> {new Date(alert.scheduledEnd).toLocaleString()}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className={styles.alertCardActions}>
                    <SharePointButton
                      variant="secondary"
                      icon={<Edit24Regular />}
                      onClick={() => handleEditAlert(alert)}
                    >
                      Edit
                    </SharePointButton>
                    <SharePointButton
                      variant="danger"
                      icon={<Delete24Regular />}
                      onClick={() => handleDeleteAlert(alert.id, alert.title)}
                    >
                      Delete
                    </SharePointButton>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Edit Alert Dialog */}
      {editingAlert && (
        <SharePointDialog
          isOpen={!!editingAlert}
          onClose={handleCancelEdit}
          title={`Edit Alert: ${editingAlert.title}`}
          width={900}
          footer={
            <div className={styles.flexRowGap12}>
              <SharePointButton
                variant="primary"
                onClick={handleSaveEdit}
                disabled={isEditingAlert}
              >
                {isEditingAlert ? 'Saving...' : 'Save Changes'}
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={handleCancelEdit}
                disabled={isEditingAlert}
              >
                Cancel
              </SharePointButton>
            </div>
          }
        >
          <div className={styles.editAlertForm}>
            <SharePointSection title="Basic Information">
              <SharePointInput
                label="Alert Title"
                value={editingAlert.title}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { ...prev, title: value } : null);
                  if (editErrors.title) setEditErrors(prev => ({ ...prev, title: undefined }));
                }}
                placeholder="Enter a clear, concise title"
                required
                error={editErrors.title}
              />

              <SharePointRichTextEditor
                label="Alert Description"
                value={editingAlert.description}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { ...prev, description: value } : null);
                  if (editErrors.description) setEditErrors(prev => ({ ...prev, description: undefined }));
                }}
                placeholder="Provide detailed information about the alert..."
                required
                error={editErrors.description}
              />
            </SharePointSection>

            <SharePointSection title="Alert Configuration">
              <SharePointSelect
                label="Alert Type"
                value={editingAlert.AlertType}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { ...prev, AlertType: value } : null);
                  if (editErrors.AlertType) setEditErrors(prev => ({ ...prev, AlertType: undefined }));
                }}
                options={alertTypeOptions}
                required
                error={editErrors.AlertType}
              />

              <SharePointSelect
                label="Priority Level"
                value={editingAlert.priority}
                onChange={(value) => setEditingAlert(prev => prev ? { ...prev, priority: value as AlertPriority } : null)}
                options={priorityOptions}
                required
              />

              <SharePointToggle
                label="Pin Alert"
                checked={editingAlert.isPinned}
                onChange={(checked) => setEditingAlert(prev => prev ? { ...prev, isPinned: checked } : null)}
              />

              <SharePointSelect
                label="Notification Type"
                value={editingAlert.notificationType}
                onChange={(value) => setEditingAlert(prev => prev ? { ...prev, notificationType: value as NotificationType } : null)}
                options={notificationOptions}
              />
            </SharePointSection>

            <SharePointSection title="Action Link (Optional)">
              <SharePointInput
                label="Link URL"
                value={editingAlert.linkUrl || ""}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { ...prev, linkUrl: value } : null);
                  if (editErrors.linkUrl) setEditErrors(prev => ({ ...prev, linkUrl: undefined }));
                }}
                placeholder="https://example.com/more-info"
                error={editErrors.linkUrl}
              />

              {editingAlert.linkUrl && (
                <SharePointInput
                  label="Link Description"
                  value={editingAlert.linkDescription || ""}
                  onChange={(value) => {
                    setEditingAlert(prev => prev ? { ...prev, linkDescription: value } : null);
                    if (editErrors.linkDescription) setEditErrors(prev => ({ ...prev, linkDescription: undefined }));
                  }}
                  placeholder="Learn More"
                  required={!!editingAlert.linkUrl}
                  error={editErrors.linkDescription}
                />
              )}
            </SharePointSection>

            <SharePointSection title="Scheduling (Optional)">
              <SharePointInput
                label="Start Date & Time"
                type="datetime-local"
                value={editingAlert.scheduledStart ? new Date(editingAlert.scheduledStart.getTime() - editingAlert.scheduledStart.getTimezoneOffset() * 60000).toISOString().slice(0, 16) : ""}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { 
                    ...prev, 
                    scheduledStart: value ? new Date(value) : undefined 
                  } : null);
                  if (editErrors.scheduledStart) setEditErrors(prev => ({ ...prev, scheduledStart: undefined }));
                }}
                error={editErrors.scheduledStart}
              />

              <SharePointInput
                label="End Date & Time"
                type="datetime-local"
                value={editingAlert.scheduledEnd ? new Date(editingAlert.scheduledEnd.getTime() - editingAlert.scheduledEnd.getTimezoneOffset() * 60000).toISOString().slice(0, 16) : ""}
                onChange={(value) => {
                  setEditingAlert(prev => prev ? { 
                    ...prev, 
                    scheduledEnd: value ? new Date(value) : undefined 
                  } : null);
                  if (editErrors.scheduledEnd) setEditErrors(prev => ({ ...prev, scheduledEnd: undefined }));
                }}
                error={editErrors.scheduledEnd}
              />
            </SharePointSection>
          </div>
        </SharePointDialog>
      )}
    </>
  );
};

export default ManageAlertsTab;