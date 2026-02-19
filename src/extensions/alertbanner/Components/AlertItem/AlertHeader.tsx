import * as React from "react";
import { IconButton } from "@fluentui/react";
import {
  ChevronDown24Regular,
  ChevronUp24Regular,
} from "@fluentui/react-icons";
import { IAlertItem } from "../Alerts/IAlerts";
import { AlertPriority } from "../Alerts/IAlerts";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { getAlertTypeIcon } from "./utils";
import styles from "./AlertItem.module.scss";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text as CoreText } from "@microsoft/sp-core-library";

interface IAlertHeaderProps {
  item: IAlertItem;
  iconName?: string;
  expanded: boolean;
  toggleExpanded: () => void;
  ariaControlsId: string;
  titleColor?: string;
}

const AlertHeader: React.FC<IAlertHeaderProps> = React.memo(
  ({
    item,
    iconName,
    expanded,
    toggleExpanded,
    ariaControlsId,
    titleColor,
  }) => {
    const priorityLabel = React.useMemo(() => {
      switch (item.priority) {
        case AlertPriority.Critical:
          return strings.PriorityCritical;
        case AlertPriority.High:
          return strings.PriorityHigh;
        case AlertPriority.Medium:
          return strings.PriorityMedium;
        case AlertPriority.Low:
        default:
          return strings.PriorityLow;
      }
    }, [item.priority]);

    const previewContent = React.useMemo(() => {
      if (!item.description) {
        return "";
      }

      return htmlSanitizer.sanitizePreviewContent(item.description);
    }, [item.description]);

    const priorityTooltip = React.useMemo(
      () => CoreText.format(strings.AlertHeaderPriorityTooltip, priorityLabel),
      [priorityLabel],
    );

    return (
      <>
        <div className={styles.iconSection}>
          <div className={styles.alertIcon} title={priorityTooltip}>
            {getAlertTypeIcon(iconName, item.priority, titleColor)}
          </div>
        </div>
        <div className={styles.textSection}>
          <div className={styles.titleRow}>
            {item.title && (
              <span
                className={styles.alertTitle}
                style={titleColor ? { color: titleColor } : undefined}
              >
                {item.title}
              </span>
            )}
            {item.AlertType && (
              <span
                className={styles.alertTypeBadge}
                style={
                  titleColor
                    ? { color: titleColor, borderColor: titleColor }
                    : undefined
                }
              >
                {item.AlertType}
              </span>
            )}
          </div>
          {!expanded && item.description && (
            <div className={styles.alertDescription} id={ariaControlsId}>
              <div
                className={styles.truncatedHtml}
                dangerouslySetInnerHTML={{
                  __html: previewContent,
                }}
              />
            </div>
          )}
        </div>
        <div className={styles.actionSection}>
          <IconButton
            onRenderIcon={() =>
              expanded ? <ChevronUp24Regular /> : <ChevronDown24Regular />
            }
            onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
              e.stopPropagation();
              toggleExpanded();
            }}
            aria-expanded={expanded}
            aria-controls={ariaControlsId}
            aria-label={expanded ? strings.CollapseAlert : strings.ExpandAlert}
            className={styles.iconButton}
          />
        </div>
      </>
    );
  },
);

AlertHeader.displayName = "AlertHeader";

export default AlertHeader;
