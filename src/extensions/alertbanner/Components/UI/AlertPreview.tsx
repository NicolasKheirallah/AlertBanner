import * as React from "react";
import { AlertPriority, IAlertType } from "../Alerts/IAlerts";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { getAlertTypeIcon } from "../AlertItem/utils";
import { getContrastText } from "../Utils/ColorUtils";
import styles from "./AlertPreview.module.scss";

const AlertPreview: React.FC<{
  title: string;
  description: string;
  alertType: IAlertType;
  priority: AlertPriority;
  isPinned?: boolean;
  className?: string;
}> = ({
  title,
  description,
  alertType,
  priority,
  isPinned = false,
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

  // Get priority-specific border color or fallback to alert type's background color
  const priorityColor = alertType.priorityColors?.[priority]?.borderColor || alertType.backgroundColor;
  const priorityTextColor = getContrastText(priorityColor);
  const normalizeColorToken = (color?: string): string =>
    (color || "").toLowerCase().replace("#", "").replace(/[^a-f0-9]/g, "");

  const colorClassMap = styles as unknown as Record<string, string>;
  const getTokenClass = (prefix: string, color: string, fallback: string): string => {
    const token = normalizeColorToken(color);
    const className =
      colorClassMap[`${prefix}${token}`] ||
      colorClassMap[`${prefix}${fallback}`];
    return className || "";
  };

  const headerBgClass = getTokenClass("previewHeaderBg", priorityColor, "0078d4");
  const headerFgClass = getTokenClass("previewHeaderFg", priorityTextColor, "ffffff");
  const backgroundSwatchClass = getTokenClass("colorSwatchBg", alertType.backgroundColor, "0078d4");
  const textSwatchClass = getTokenClass("colorSwatchText", alertType.textColor, "323130");

  return (
    <div className={`${styles.previewContainer} ${className || ''}`}>
      <div className={`${styles.previewHeader} ${styles.previewHeaderDynamic} ${headerBgClass} ${headerFgClass}`}>
        <h4 className={styles.headerTitle}>Live Preview</h4>
        <span className={styles.previewNote}>This is how your alert will appear</span>
      </div>

      <div
        className={`${styles.alertPreview} ${getPriorityClass(priority)} ${isPinned ? styles.pinned : ''}`}
        style={{
          borderLeftColor: priorityColor,
        } as React.CSSProperties}
      >
        <div className={styles.headerRow}>
          <div className={styles.iconSection}>
            <div className={styles.alertIcon}>
              {getAlertTypeIcon(alertType.iconName, priority)}
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
                <div dangerouslySetInnerHTML={{ 
                  __html: React.useMemo(() => 
                    htmlSanitizer.sanitizePreviewContent(description || 'Alert description will appear here...'), 
                    [description]
                  )
                }} />
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
              <span className={`${styles.colorSwatch} ${styles.colorSwatchBackground} ${backgroundSwatchClass}`} />
              {alertType.backgroundColor}
            </span>
          </div>
          <div className={styles.infoItem}>
            <span className={styles.infoLabel}>Text Color:</span>
            <span className={styles.infoValue}>
              <span className={`${styles.colorSwatch} ${styles.colorSwatchText} ${textSwatchClass}`} />
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
