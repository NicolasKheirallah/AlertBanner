import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import { ChevronDown24Regular, ChevronUp24Regular } from "@fluentui/react-icons";
import { IAlertItem } from "../Alerts/IAlerts";
import { AlertPriority } from "../Alerts/IAlerts";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { getPriorityIcon } from "./utils";
import styles from "./AlertItem.module.scss";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';

interface IAlertHeaderProps {
  item: IAlertItem;
  expanded: boolean;
  toggleExpanded: () => void;
  ariaControlsId: string;
}

const AlertHeader: React.FC<IAlertHeaderProps> = React.memo(({ item, expanded, toggleExpanded, ariaControlsId }) => {
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
      return '';
    }

    return htmlSanitizer.sanitizePreviewContent(item.description);
  }, [item.description]);

  const priorityTooltip = React.useMemo(
    () => CoreText.format(strings.AlertHeaderPriorityTooltip, priorityLabel),
    [priorityLabel]
  );

  return (
    <>
      <div className={styles.iconSection}>
        <div className={styles.alertIcon} title={priorityTooltip}>
          {getPriorityIcon(item.priority)}
        </div>
      </div>
      <div className={styles.textSection}>
        {item.title && (
          <Text className={styles.alertTitle} size={500} weight="semibold">
            {item.title}
          </Text>
        )}
        {!expanded && item.description && (
          <div className={styles.alertDescription} id={ariaControlsId}>
            <div
              className={styles.truncatedHtml}
              dangerouslySetInnerHTML={{
                __html: previewContent
              }}
            />
          </div>
        )}
      </div>
      <div className={styles.actionSection}>
        <Button
          appearance="subtle"
          icon={expanded ? <ChevronUp24Regular /> : <ChevronDown24Regular />}
          onClick={(e: React.MouseEvent) => {
            e.stopPropagation();
            toggleExpanded();
          }}
          aria-expanded={expanded}
          aria-controls={ariaControlsId}
          aria-label={expanded ? strings.CollapseAlert : strings.ExpandAlert}
          size="small"
        />
      </div>
    </>
  );
});

AlertHeader.displayName = 'AlertHeader';

export default AlertHeader;
