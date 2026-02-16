import * as React from "react";
import { AlertPriority, IAlertType } from "../Alerts/IAlerts";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { getAlertTypeIcon } from "../AlertItem/utils";
import { getContrastText } from "../Utils/ColorUtils";
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

  // Always use white background for preview with dark text
  const containerStyle: React.CSSProperties = {
    backgroundColor: '#ffffff',
    color: '#323130', // Dark grey text
    border: '1px solid #edebe9',
  };

  const textStyle: React.CSSProperties = {
    color: '#323130', // Dark grey text
    textShadow: 'none',
    WebkitFontSmoothing: 'antialiased',
    MozOsxFontSmoothing: 'grayscale',
  };

  // Use the alert type's background color for priority indicators
  const priorityColor = alertType.backgroundColor;
  const priorityTextColor = getContrastText(priorityColor);

  // Header style - uses alert type background color
  const headerStyle: React.CSSProperties = {
    backgroundColor: priorityColor,
    color: priorityTextColor,
    border: '1px solid #edebe9',
  };

  // Container style with left border showing priority color
  const alertContainerStyle: React.CSSProperties = {
    ...containerStyle,
    borderLeft: `4px solid ${priorityColor}`,
  };

  return (
    <div className={`${styles.previewContainer} ${className || ''}`}>
      <div className={styles.previewHeader} style={headerStyle}>
        <h4 style={{ color: priorityTextColor }}>Live Preview</h4>
        <span className={styles.previewNote} style={{ color: priorityTextColor, opacity: 0.9 }}>This is how your alert will appear</span>
      </div>

      <div
        className={`${styles.alertPreview} ${getPriorityClass(priority)} ${isPinned ? styles.pinned : ''}`}
        style={alertContainerStyle}
      >
        <div className={styles.headerRow}>
          <div className={styles.iconSection}>
            <div className={styles.alertIcon}>
              {getAlertTypeIcon(alertType.iconName, priority)}
            </div>
          </div>

          <div className={styles.textSection}>
            {title && (
              <div className={styles.alertTitle} style={textStyle}>
                {title || 'Alert Title'}
                {isPinned && <span className={styles.pinnedBadge} style={textStyle}>📌 PINNED</span>}
              </div>
            )}

            {description && (
              <div className={styles.alertDescription} style={textStyle}>
                <div dangerouslySetInnerHTML={{ 
                  __html: React.useMemo(() => 
                    htmlSanitizer.sanitizePreviewContent(description || 'Alert description will appear here...'), 
                    [description]
                  )
                }} />
              </div>
            )}

            {linkUrl && linkDescription && (
              <div className={styles.alertLink}>
                <a href={linkUrl} target="_blank" rel="noopener noreferrer" style={textStyle}>
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
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Icon:</span>
            <span className={styles.infoValue}>{alertType.iconName}</span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AlertPreview;
