import * as React from "react";
import {
  IStackItemStyles,
  IconButton,
  Stack,
  ResponsiveMode,
} from "office-ui-fabric-react";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { css } from "@uifabric/utilities/lib/css";
import { AlertType, IAlertItem } from "../Alerts/index";
import styles from "./AlertItem.module.scss";

export interface IAlertItemProps {
  item: IAlertItem;
  remove: (id: number) => void;
  responsiveMode?: ResponsiveMode;
}

// Stack item styles for consistent layout
const stackItemStyles: IStackItemStyles = {
  root: {
    display: "flex",
    justifyContent: "inherit",
    overflow: "hidden",
  },
};

// Mapping of alert types to corresponding styles and icons
const alertTypeMapping: Record<AlertType, { style: string; icon: string }> = {
  [AlertType.Interruption]: { style: styles.actionable, icon: "IncidentTriangle" },
  [AlertType.Warning]: { style: styles.warning, icon: "ShieldAlert" },
  [AlertType.Maintenance]: { style: styles.maintenance, icon: "CRMServices" },
  [AlertType.Info]: { style: styles.info, icon: "Info12" },
};

// Type guard to check if a value is a valid AlertType
const isAlertType = (type: string): type is AlertType => {
  return Object.keys(AlertType).includes(type);
};

// Optimized AlertItem component
const AlertItem: React.FC<IAlertItemProps> = React.memo(({ item, remove, responsiveMode }) => {
  const [expanded, setExpanded] = React.useState(false); // Track expanded state
  const nodeRef = React.useRef<HTMLDivElement | null>(null); // Reference to the DOM node

  // Collapse alert if clicked outside of it
  React.useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      if (nodeRef.current && !nodeRef.current.contains(event.target as Node)) {
        setExpanded(false);
      }
    };
    document.addEventListener("mousedown", handleOutsideClick);
    return () => {
      document.removeEventListener("mousedown", handleOutsideClick);
    };
  }, []);

  // Memoized toggle function to prevent unnecessary re-renders
  const toggleExpanded = React.useCallback(() => {
    setExpanded((prev) => !prev);
  }, []);

  // Memoized remove function to prevent unnecessary re-renders
  const handleRemove = React.useCallback(() => {
    remove(item.Id);
  }, [remove, item.Id]);

  // Memoize the alert type styling and icon to avoid recalculation
  const { style, icon } = React.useMemo(() => {
    if (!isAlertType(item.AlertType)) {
      console.error(`Unexpected alert type: ${item.AlertType}`);
      return { style: "", icon: "" }; // Return empty strings if the alert type is invalid
    }
    return alertTypeMapping[item.AlertType];
  }, [item.AlertType]);

  // Memoize the description class based on expanded state
  const descriptionClassName = React.useMemo(() => {
    return expanded ? styles.alertDescriptionExp : styles.alertDescription;
  }, [expanded]);

  // Render a stack item with specified content, class, and grow properties
  const renderStackItem = React.useCallback(
    (content: React.ReactNode, className: string = "", grow: number = 1): JSX.Element => (
      <Stack.Item align="center" grow={grow} styles={stackItemStyles} className={className}>
        {content}
      </Stack.Item>
    ),
    []
  );

  // Render an icon button with a specified icon and click handler
  const renderIconButton = React.useCallback(
    (iconName: string, onClick: () => void): JSX.Element => (
      <IconButton onClick={onClick}>
        <Icon iconName={iconName} />
      </IconButton>
    ),
    []
  );

  // Render the link associated with the alert
  const renderLink = React.useCallback((): JSX.Element | null => {
    if (!item.link) return null;
    return renderStackItem(
      <span className={styles.alertLink}>
        <a href={item.link.Url} title={item.link.Description}>
          {item.link.Description}
        </a>
      </span>,
      css(styles.stackLink, styles.hideInMobile),
      0 // No growth for the link
    );
  }, [item.link, renderStackItem]);

  return (
    <div className={styles.alertItem} ref={nodeRef}>
      <div className={css(styles.container, style)}>
        <Stack horizontal className={styles.stack}>
          {renderStackItem(
            <span className={styles.alertIcon}>
              <Icon iconName={icon} />
            </span>,
            "",
            0 // Icon should not grow
          )}
          {renderStackItem(
            <div>
              {item.title && (
                <span className={styles.alertTitle}>{item.title}</span>
              )}
              {item.description && (
                <span
                  className={descriptionClassName}
                  dangerouslySetInnerHTML={{ __html: item.description }}
                />
              )}
              {expanded && item.link && (
                <span className={css(styles.alertLink)}>
                  <a href={item.link.Url}></a>
                </span>
              )}
            </div>,
            styles.stackMessage,
            1 // Allow the message to grow
          )}
          {responsiveMode === ResponsiveMode.large || responsiveMode === ResponsiveMode.xLarge ? (
            <>
              {renderStackItem(
                <span className={styles.alertClose}>
                  {renderIconButton(
                    expanded ? "ChevronUp" : "ChevronDown",
                    toggleExpanded
                  )}
                </span>,
                styles.hideInMobile,
                0 // Close button should not grow
              )}
              {renderStackItem(
                <span className={styles.alertClose}>
                  {expanded
                    ? renderIconButton("ChromeClose", handleRemove)
                    : renderIconButton("ChevronDown", toggleExpanded)}
                </span>,
                styles.hideInDesktop,
                0 // Close button should not grow
              )}
              {renderLink()}
              {renderStackItem(
                <span className={styles.alertClose}>
                  {renderIconButton("ChromeClose", handleRemove)}
                </span>,
                styles.hideInMobile,
                0 // Close button should not grow
              )}
            </>
          ) : (
            <>
              {renderStackItem(
                <span className={styles.alertClose}>
                  {renderIconButton(
                    expanded ? "ChevronUp" : "ChevronDown",
                    toggleExpanded
                  )}
                </span>,
                "",
                0 // Close button should not grow
              )}
              {expanded && renderLink()}
              {renderStackItem(
                <span className={styles.alertClose}>
                  {renderIconButton("ChromeClose", handleRemove)}
                </span>,
                "",
                0 // Close button should not grow
              )}
            </>
          )}
        </Stack>
      </div>
    </div>
  );
});

export default AlertItem;
