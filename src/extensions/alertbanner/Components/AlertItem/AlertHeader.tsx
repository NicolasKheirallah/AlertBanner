import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import { ChevronDown24Regular, ChevronUp24Regular } from "@fluentui/react-icons";
import { IAlertItem } from "../Alerts/IAlerts";
import { getPriorityIcon } from "./utils"; // getPriorityIcon is exported from utils.tsx
import styles from "./AlertItem.module.scss";
// DescriptionContent is not rendered directly in AlertHeader, it's in AlertContent

interface IAlertHeaderProps {
  item: IAlertItem;
  expanded: boolean;
  toggleExpanded: () => void;
  ariaControlsId: string;
}

const AlertHeader: React.FC<IAlertHeaderProps> = React.memo(({ item, expanded, toggleExpanded, ariaControlsId }) => {
  return (
    <>
      <div className={styles.iconSection}>
        <div className={styles.alertIcon}>
          {getPriorityIcon(item.priority)}
        </div>
      </div>
      <div className={styles.textSection}>
        {item.title && (
          <Text className={styles.alertTitle} size={500} weight="semibold">
            {item.title}
          </Text>
        )}
        {/* Description is handled by AlertContent, not AlertHeader */}
      </div>
      <div className={styles.actionSection}>
        <Button
          appearance="transparent"
          icon={expanded ? <ChevronUp24Regular /> : <ChevronDown24Regular />}
          onClick={toggleExpanded}
          aria-expanded={expanded}
          aria-controls={ariaControlsId}
          aria-label={expanded ? "Collapse Alert" : "Expand Alert"}
          size="small"
        />
      </div>
    </>
  );
});

export default AlertHeader;
