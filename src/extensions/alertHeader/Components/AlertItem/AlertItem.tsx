import * as React from "react";
import {
  IconButton,
  ResponsiveMode,
} from "office-ui-fabric-react";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { css } from "@uifabric/utilities/lib/css";
import { AlertType, IAlertItem } from "../Alerts/IAlerts.types";
import styles from "./AlertItem.module.scss";

export interface IAlertItemProps {
  item: IAlertItem;
  remove: (id: number) => void;
  responsiveMode?: ResponsiveMode;
}

const AlertItem: React.FC<IAlertItemProps> = ({
  item,
  remove,
  responsiveMode,
}) => {
  const [expanded, setExpanded] = React.useState(false);
  const nodeRef = React.useRef<HTMLDivElement | null>(null);
  const ariaControlsId = `alert-description-${item.Id}`;

  const toggleExpanded = () => {
    setExpanded((prev) => !prev);
  };

  const handleRemove = () => {
    remove(item.Id);
  };

  React.useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      if (
        nodeRef.current &&
        !nodeRef.current.contains(event.target as Node)
      ) {
        setExpanded(false);
      }
    };
    document.addEventListener("mousedown", handleOutsideClick);
    return () => {
      document.removeEventListener("mousedown", handleOutsideClick);
    };
  }, []);

  const { style: alertStyle, icon } = alertTypeMapping[item.AlertType];

  const descriptionClassName = expanded
    ? styles.alertDescriptionExp
    : styles.alertDescription;

  // Render the link associated with the alert
  const renderLink = () => {
    if (!item.link) return null;
    return (
      <div className={styles.alertLink}>
        <a href={item.link.Url} title={item.link.Description}>
          {item.link.Description}
        </a>
      </div>
    );
  };

  return (
    <div className={styles.alertItem} ref={nodeRef}>
      <div className={css(styles.container, alertStyle)}>
        {/* Icon Section */}
        <div className={styles.iconSection}>
          <Icon iconName={icon} className={styles.alertIcon} />
        </div>

        {/* Text Section */}
        <div className={styles.textSection}>
          {item.title && (
            <div className={styles.alertTitle}>{item.title}</div>
          )}
          {item.description && (
            <div
              className={descriptionClassName}
              id={ariaControlsId}
              dangerouslySetInnerHTML={{ __html: item.description }}
            ></div>
          )}
        </div>

        {/* Action Section */}
        <div className={styles.actionSection}>
          {/* Render link if expanded */}
          {renderLink()}
          <IconButton
            onClick={toggleExpanded}
            aria-expanded={expanded}
            aria-controls={ariaControlsId}
            aria-label={expanded ? "Collapse Alert" : "Expand Alert"}
            className={styles.toggleButton}
          >
            <Icon iconName={expanded ? "ChevronUp" : "ChevronDown"} />
          </IconButton>
          <IconButton
            onClick={handleRemove}
            aria-label="Close Alert"
            className={styles.closeButton}
          >
            <Icon iconName="ChromeClose" />
          </IconButton>
        </div>
      </div>
    </div>
  );
};

// Mapping of alert types to corresponding styles and icons
const alertTypeMapping: Record<AlertType, { style: string; icon: string }> = {
  [AlertType.Interruption]: {
    style: styles.interruption,
    icon: "IncidentTriangle",
  },
  [AlertType.Warning]: { style: styles.warning, icon: "ShieldAlert" },
  [AlertType.Maintenance]: { style: styles.maintenance, icon: "CRMServices" },
  [AlertType.Info]: { style: styles.info, icon: "Info12" },
};

export default AlertItem;
