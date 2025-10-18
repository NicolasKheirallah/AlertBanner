import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import { ChevronDown24Regular, ChevronUp24Regular } from "@fluentui/react-icons";
import { IAlertItem } from "../Services/SharePointAlertService";
import { AlertPriority } from "../Alerts/IAlerts";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { getPriorityIcon } from "./utils";
import { useLocalization } from "../Hooks/useLocalization";
import styles from "./AlertItem.module.scss";

interface IAlertHeaderProps {
  item: IAlertItem;
  expanded: boolean;
  toggleExpanded: () => void;
  ariaControlsId: string;
}

const AlertHeader: React.FC<IAlertHeaderProps> = React.memo(({ item, expanded, toggleExpanded, ariaControlsId }) => {
  const { getString } = useLocalization();

  const priorityLabel = React.useMemo(() => {
    switch (item.priority) {
      case AlertPriority.Critical:
        return getString('PriorityCritical');
      case AlertPriority.High:
        return getString('PriorityHigh');
      case AlertPriority.Medium:
        return getString('PriorityMedium');
      case AlertPriority.Low:
      default:
        return getString('PriorityLow');
    }
  }, [getString, item.priority]);

  const previewContent = React.useMemo(() => {
    if (!item.description) {
      return '';
    }

    return htmlSanitizer.sanitizePreviewContent(item.description);
  }, [item.description]);

  const priorityTooltip = React.useMemo(
    () => getString('AlertHeaderPriorityTooltip', priorityLabel),
    [getString, priorityLabel]
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
          aria-label={expanded ? getString('CollapseAlert') : getString('ExpandAlert')}
          size="small"
        />
      </div>
    </>
  );
});

AlertHeader.displayName = 'AlertHeader';

export default AlertHeader;
