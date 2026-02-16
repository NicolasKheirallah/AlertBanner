import * as React from "react";
import { logger } from "../Services/LoggerService";
import { IAlertType, IQuickAction } from "../Alerts/IAlerts";
import { IAlertItem } from "../Alerts/IAlerts";
import { getContrastText } from "../Utils/ColorUtils";
import styles from "./AlertItem.module.scss";
import AlertHeader from "./AlertHeader";
import AlertContent from "./AlertContent";
import AlertActions from "./AlertActions";
import { WINDOW_OPEN_CONFIG, SHADOW_CONFIG } from "../Utils/AppConstants";

// Whitelist of safe CSS properties for alert styling
const ALLOWED_CSS_PROPERTIES = new Set([
  "color",
  "backgroundColor",
  "background",
  "border",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "borderRadius",
  "padding",
  "margin",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "textAlign",
  "textDecoration",
  "boxShadow",
  "opacity",
  "display",
  "width",
  "height",
  "maxWidth",
  "minWidth",
  "maxHeight",
  "minHeight",
  "lineHeight",
  "letterSpacing",
]);

// Dangerous patterns that could enable CSS injection attacks
const DANGEROUS_CSS_PATTERNS = [
  /url\s*\(/i, // Block url() - potential data exfiltration
  /expression\s*\(/i, // Block expression() - IE XSS vector
  /javascript\s*:/i, // Block javascript: protocol
  /data\s*:/i, // Block data: URIs
  /behavior\s*:/i, // Block behavior: - IE XSS vector
  /@import/i, // Block @import
  /binding\s*:/i, // Block -moz-binding
  /\\[0-9a-f]/i, // Block unicode escape sequences
];

const parseAdditionalStyles = (stylesString?: string): React.CSSProperties => {
  if (!stylesString) return {};

  const styleObj: Record<string, string | number> = {};
  const stylesArray = stylesString.split(";").filter((s) => s.trim());

  stylesArray.forEach((style) => {
    const colonIndex = style.indexOf(":");
    if (colonIndex === -1) return;

    const key = style.substring(0, colonIndex).trim();
    const value = style.substring(colonIndex + 1).trim();

    if (!key || !value) return;

    const camelCaseKey = key.replace(/-([a-z])/g, (_, group1) =>
      group1.toUpperCase(),
    );

    // Security: Only allow whitelisted CSS properties
    if (!ALLOWED_CSS_PROPERTIES.has(camelCaseKey)) {
      logger.warn("AlertItem", `Blocked non-whitelisted CSS property: ${key}`);
      return;
    }

    // Security: Block dangerous CSS values
    for (const pattern of DANGEROUS_CSS_PATTERNS) {
      if (pattern.test(value)) {
        logger.warn(
          "AlertItem",
          `Blocked dangerous CSS value: ${value.substring(0, 50)}`,
        );
        return;
      }
    }

    styleObj[camelCaseKey] = isNaN(Number(value)) ? value : Number(value);
  });

  return styleObj as React.CSSProperties;
};

export interface IAlertItemProps {
  item: IAlertItem;
  remove: (id: string) => void;
  hideForever: (id: string) => void;
  alertType: IAlertType;
  isCarousel?: boolean;
  currentIndex?: number;
  totalAlerts?: number;
  onNext?: () => void;
  onPrevious?: () => void;
  userTargetingEnabled?: boolean;
}

const AlertItem: React.FC<IAlertItemProps> = ({
  item,
  remove,
  hideForever,
  alertType,
  isCarousel = false,
  currentIndex = 1,
  totalAlerts = 1,
  onNext,
  onPrevious,
  userTargetingEnabled = false,
}) => {
  const [expanded, setExpanded] = React.useState(false);
  const ariaControlsId = `alert-description-${item.id}`;

  const handlers = React.useMemo(
    () => ({
      toggleExpanded: () => setExpanded((prev) => !prev),
      remove: (id: string) => remove(id),
      hideForever: (id: string) => hideForever(id),
      stopPropagation: (e: React.MouseEvent) => e.stopPropagation(),
    }),
    [remove, hideForever],
  );

  const handleQuickAction = React.useCallback(
    (action: IQuickAction) => {
      switch (action.actionType) {
        case "link":
          if (action.url) {
            window.open(
              action.url,
              WINDOW_OPEN_CONFIG.TARGET,
              WINDOW_OPEN_CONFIG.FEATURES,
            );
          }
          break;
        case "dismiss":
          handlers.remove(item.id);
          break;
        case "acknowledge":
          logger.debug("AlertItem", `Alert ${item.id} acknowledged`);
          handlers.remove(item.id);
          break;
        case "custom":
          const allowedCustomActions: {
            [key: string]: (item: IAlertItem) => void;
          } = {
            showDetails: (item) => {
              logger.debug(
                "AlertItem",
                `Showing details for alert: ${item.id}`,
              );
            },
            logInteraction: (item) => {
              logger.debug(
                "AlertItem",
                `User interacted with alert: ${item.id}`,
              );
            },
            markAsRead: (item) => {
              logger.debug("AlertItem", `Marking alert as read: ${item.id}`);
              handlers.remove(item.id);
            },
          };

          if (
            action.callback &&
            typeof allowedCustomActions[action.callback] === "function"
          ) {
            allowedCustomActions[action.callback](item);
          } else {
            logger.warn(
              "AlertItem",
              `Unknown or disallowed custom action: ${action.callback}. Allowed actions: ${Object.keys(allowedCustomActions).join(", ")}`,
            );
          }
          break;
      }
    },
    [handlers.remove, item],
  );

  const baseContainerStyle = React.useMemo<React.CSSProperties>(() => {
    const backgroundColor = alertType.backgroundColor || "#389899";
    const textColor = alertType.textColor || getContrastText(backgroundColor);

    return {
      backgroundColor,
      color: textColor,
      ...parseAdditionalStyles(alertType.additionalStyles),
    };
  }, [alertType]);

  const priorityStyle = React.useMemo(
    () =>
      alertType.priorityStyles
        ? alertType.priorityStyles[
            item.priority as keyof typeof alertType.priorityStyles
          ]
        : "",
    [alertType.priorityStyles, item.priority],
  );

  const containerStyle = React.useMemo<React.CSSProperties>(
    () => ({
      ...baseContainerStyle,
      ...parseAdditionalStyles(priorityStyle),
      ...(item.priority === "critical" && {
        boxShadow: SHADOW_CONFIG.CRITICAL_PRIORITY,
      }),
    }),
    [baseContainerStyle, priorityStyle, item.priority],
  );

  const containerClassNames = [
    styles.container,
    styles.clickable,
    item.priority === "critical" ? styles.critical : "",
    item.priority === "high" ? styles.high : "",
    item.priority === "medium" ? styles.medium : "",
    item.priority === "low" ? styles.low : "",
    item.isPinned ? styles.pinned : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div className={styles.alertItem}>
      <div className={containerClassNames} style={containerStyle}>
        <div
          className={styles.headerRow}
          onClick={handlers.toggleExpanded}
          role="button"
          tabIndex={0}
          aria-expanded={expanded}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              handlers.toggleExpanded();
            }
          }}
        >
          <AlertHeader
            item={item}
            iconName={alertType.iconName}
            expanded={expanded}
            toggleExpanded={handlers.toggleExpanded}
            ariaControlsId={ariaControlsId}
          />
          <AlertActions
            item={item}
            isCarousel={isCarousel}
            currentIndex={currentIndex}
            totalAlerts={totalAlerts}
            onNext={onNext}
            onPrevious={onPrevious}
            expanded={expanded}
            toggleExpanded={handlers.toggleExpanded}
            remove={handlers.remove}
            hideForever={handlers.hideForever}
            stopPropagation={handlers.stopPropagation}
            userTargetingEnabled={userTargetingEnabled}
          />
        </div>
        <AlertContent
          item={item}
          expanded={expanded}
          stopPropagation={handlers.stopPropagation}
          ariaControlsId={ariaControlsId}
        />
      </div>
    </div>
  );
};

export default AlertItem;
