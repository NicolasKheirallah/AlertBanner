import * as React from "react";
import { AlertPriority, IAlertType } from "../Alerts/IAlerts";
import { getPriorityIcon } from "../AlertItem/utils";
import styles from "./AlertPreview.module.scss";

interface IAlertPreviewProps {
  title: string;
  description: string;
  alertType: IAlertType;
  priority: AlertPriority;
  isPinned?: boolean;
  linkUrl?: string;
  linkDescription?: string;
  className?: string;
}

const AlertPreview: React.FC<IAlertPreviewProps> = ({
  title,
  description,
  alertType,
  priority,
  isPinned = false,
  linkUrl,
  linkDescription,
  className
}) => {
  const getPriorityClass = (priority: AlertPriority): string => {
    switch (priority) {
      case AlertPriority.Critical: return styles.critical;
      case AlertPriority.High: return styles.high;
      case AlertPriority.Medium: return styles.medium;
      case AlertPriority.Low: return styles.low;
      default: return styles.medium;
    }
  };

  // Ensure proper contrast for preview
  const getContrastText = (bgColor: string): string => {
    // Simple contrast check - if background is very light, use dark text
    if (bgColor.toLowerCase() === '#ffffff' || bgColor.toLowerCase() === 'white') {
      return '#323130'; // Dark text for white background
    }
    return alertType.textColor;
  };

  const containerStyle: React.CSSProperties = {
    backgroundColor: alertType.backgroundColor,
    color: getContrastText(alertType.backgroundColor),
    border: alertType.backgroundColor === '#ffffff' ? '1px solid #edebe9' : undefined,
  };

  return (
    <div className={`${styles.previewContainer} ${className || ''}`}>
      <div className={styles.previewHeader}>
        <h4>Live Preview</h4>
        <span className={styles.previewNote}>This is how your alert will appear</span>
      </div>
      
      <div 
        className={`${styles.alertPreview} ${getPriorityClass(priority)} ${isPinned ? styles.pinned : ''}`}
        style={containerStyle}
      >
        <div className={styles.headerRow}>
          <div className={styles.iconSection}>
            <div className={styles.alertIcon}>
              {getPriorityIcon(priority)}
            </div>
          </div>
          
          <div className={styles.textSection}>
            {title && (
              <div className={styles.alertTitle}>
                {title || 'Alert Title'}
                {isPinned && <span className={styles.pinnedBadge}>📌 PINNED</span>}
              </div>
            )}
            
            {description && (
              <div className={styles.alertDescription}>
                <div dangerouslySetInnerHTML={{ __html: description || 'Alert description will appear here...' }} />
              </div>
            )}
            
            {linkUrl && linkDescription && (
              <div className={styles.alertLink}>
                <a href={linkUrl} target="_blank" rel="noopener noreferrer">
                  🔗 {linkDescription}
                </a>
              </div>
            )}
          </div>
          
          <div className={styles.actionSection}>
            <button className={styles.expandButton} type="button">
              ⌄
            </button>
            <button className={styles.dismissButton} type="button">
              ✕
            </button>
          </div>
        </div>
      </div>
      
      <div className={styles.previewInfo}>
        <div className={styles.infoGrid}>
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Type:</span>
            <span className={styles.infoValue}>{alertType.name}</span>
          </div>
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Priority:</span>
            <span className={`${styles.infoValue} ${styles.priorityBadge} ${getPriorityClass(priority)}`}>
              {priority.toUpperCase()}
            </span>
          </div>
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Background:</span>
            <span className={styles.infoValue}>
              <span className={styles.colorSwatch} style={{ backgroundColor: alertType.backgroundColor }} />
              {alertType.backgroundColor}
            </span>
          </div>
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Text Color:</span>
            <span className={styles.infoValue}>
              <span className={styles.colorSwatch} style={{ backgroundColor: alertType.textColor }} />
              {alertType.textColor}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AlertPreview;