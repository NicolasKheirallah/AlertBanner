import * as React from "react";
import { logger } from "../Services/LoggerService";
import { IAlertType, IAlertItem, AlertPriority } from "../Alerts/IAlerts";
import { useAlerts } from "../Context/AlertsContext";
import { getContrastText } from "../Utils/ColorUtils";
import styles from "./AlertItem.module.scss";
import AlertHeader from "./AlertHeader";
import AlertContent from "./AlertContent";
import AlertActions from "./AlertActions";
import { SHADOW_CONFIG } from "../Utils/AppConstants";

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

  // Memoized keyboard event handler for accessibility
  const handleHeaderKeyDown = React.useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handlers.toggleExpanded();
      }
    },
    [handlers.toggleExpanded],
  );

  // Memoized header click handler
  const handleHeaderClick = React.useCallback(() => {
    handlers.toggleExpanded();
  }, [handlers.toggleExpanded]);

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

  // Get global priority border colors from context
  const { state: alertsState } = useAlerts();
  const priorityBorderColors = alertsState.priorityBorderColors;
  
  // Get border color from global settings based on alert priority
  const borderColor = React.useMemo(() => {
    return priorityBorderColors[item.priority as AlertPriority]?.borderColor || 
           alertType.backgroundColor || 
           "#389899";
  }, [priorityBorderColors, item.priority, alertType.backgroundColor]);

  const containerStyle = React.useMemo<React.CSSProperties>(
    () => ({
      ...baseContainerStyle,
      ...parseAdditionalStyles(priorityStyle),
      // Set border-left color directly using priority-specific or alert type background color
      borderLeft: `4px solid ${borderColor}`,
      ...(item.priority === "critical" && {
        boxShadow: SHADOW_CONFIG.CRITICAL_PRIORITY,
      }),
    } as React.CSSProperties),
    [baseContainerStyle, priorityStyle, item.priority, borderColor],
  );

  const containerClassNames = React.useMemo(
    () =>
      [
        styles.container,
        styles.clickable,
        item.priority === "critical" ? styles.critical : "",
        item.priority === "high" ? styles.high : "",
        item.priority === "medium" ? styles.medium : "",
        item.priority === "low" ? styles.low : "",
        item.isPinned ? styles.pinned : "",
      ]
        .filter(Boolean)
        .join(" "),
    [item.priority, item.isPinned],
  );

  return (
    <div className={styles.alertItem}>
      <div className={containerClassNames} style={containerStyle}>
        <div
          className={styles.headerRow}
          onClick={handleHeaderClick}
          role="button"
          tabIndex={0}
          aria-expanded={expanded}
          onKeyDown={handleHeaderKeyDown}
        >
          <AlertHeader
            item={item}
            iconName={alertType.iconName}
            expanded={expanded}
            toggleExpanded={handlers.toggleExpanded}
            ariaControlsId={ariaControlsId}
            titleColor={undefined}
          />
          <AlertActions
            item={item}
            isCarousel={isCarousel}
            currentIndex={currentIndex}
            totalAlerts={totalAlerts}
            onNext={onNext}
            onPrevious={onPrevious}
            expanded={expanded}
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

// Custom comparison function for React.memo
// Only re-render when alert data or relevant props actually change
const arePropsEqual = (
  prevProps: IAlertItemProps,
  nextProps: IAlertItemProps,
): boolean => {
  // Compare alert item data
  if (prevProps.item.id !== nextProps.item.id) return false;
  if (prevProps.item.title !== nextProps.item.title) return false;
  if (prevProps.item.description !== nextProps.item.description) return false;
  if (prevProps.item.priority !== nextProps.item.priority) return false;
  if (prevProps.item.isPinned !== nextProps.item.isPinned) return false;
  if (prevProps.item.linkUrl !== nextProps.item.linkUrl) return false;
  if (prevProps.item.AlertType !== nextProps.item.AlertType) return false;

  // Compare carousel-related props
  if (prevProps.isCarousel !== nextProps.isCarousel) return false;
  if (prevProps.currentIndex !== nextProps.currentIndex) return false;
  if (prevProps.totalAlerts !== nextProps.totalAlerts) return false;

  // Compare alert type
  if (prevProps.alertType !== nextProps.alertType) return false;

  // Compare user targeting
  if (prevProps.userTargetingEnabled !== nextProps.userTargetingEnabled) return false;

  // Compare callbacks (reference equality)
  if (prevProps.remove !== nextProps.remove) return false;
  if (prevProps.hideForever !== nextProps.hideForever) return false;
  if (prevProps.onNext !== nextProps.onNext) return false;
  if (prevProps.onPrevious !== nextProps.onPrevious) return false;

  return true;
};

export default React.memo(AlertItem, arePropsEqual);
