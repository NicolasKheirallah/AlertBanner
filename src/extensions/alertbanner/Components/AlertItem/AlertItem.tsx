import * as React from "react";
import { logger } from "../Services/LoggerService";
import { IAlertType, IAlertItem, AlertPriority } from "../Alerts/IAlerts";
import { useAlertsState } from "../Context/AlertsContext";
import { getContrastText } from "../Utils/ColorUtils";
import styles from "./AlertItem.module.scss";
import AlertHeader from "./AlertHeader";
import AlertContent from "./AlertContent";
import AlertActions from "./AlertActions";
import { SHADOW_CONFIG } from "../Utils/AppConstants";

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

    if (!ALLOWED_CSS_PROPERTIES.has(camelCaseKey)) {
      logger.warn("AlertItem", `Blocked non-whitelisted CSS property: ${key}`);
      return;
    }

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

  const handleHeaderKeyDown = React.useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        handlers.toggleExpanded();
      }
    },
    [handlers.toggleExpanded],
  );

  const handleHeaderClick = React.useCallback(() => {
    handlers.toggleExpanded();
  }, [handlers.toggleExpanded]);

  const baseContainerStyle = React.useMemo<React.CSSProperties>(() => {
    const backgroundColor = "#ffffff";
    const textColor = "#323130";

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

  const alertsState = useAlertsState();
  const priorityBorderColors = alertsState.priorityBorderColors;

  const extractBorderColor = (borderStr?: string): string | null => {
    if (!borderStr) return null;
    const match = borderStr.match(/#[a-fA-F0-9]{6}|#[a-fA-F0-9]{3}/);
    return match ? match[0] : null;
  };
  const borderColor = React.useMemo(() => {
    const priorityStyleColor = extractBorderColor(priorityStyle);
    return (
      priorityStyleColor ||
      alertType.backgroundColor ||
      priorityBorderColors[item.priority as AlertPriority]?.borderColor ||
      "#389899"
    );
  }, [
    priorityBorderColors,
    item.priority,
    alertType.backgroundColor,
    priorityStyle,
  ]);

  const containerStyle = React.useMemo<React.CSSProperties>(() => {
    const parsedPriorityStyle = parseAdditionalStyles(priorityStyle);

    const hasPriorityBorder =
      parsedPriorityStyle.border ||
      parsedPriorityStyle.borderLeft ||
      parsedPriorityStyle.borderWidth;

    return {
      ...baseContainerStyle,
      ...parsedPriorityStyle,
      ...(!hasPriorityBorder && {
        borderLeft: `4px solid ${borderColor}`,
      }),
      ...(item.priority === "critical" && {
        boxShadow: SHADOW_CONFIG.CRITICAL_PRIORITY,
      }),
    } as React.CSSProperties;
  }, [baseContainerStyle, priorityStyle, item.priority, borderColor]);

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
            titleColor={borderColor}
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

const arePropsEqual = (
  prevProps: IAlertItemProps,
  nextProps: IAlertItemProps,
): boolean => {
  if (prevProps.item.id !== nextProps.item.id) return false;
  if (prevProps.item.title !== nextProps.item.title) return false;
  if (prevProps.item.description !== nextProps.item.description) return false;
  if (prevProps.item.priority !== nextProps.item.priority) return false;
  if (prevProps.item.isPinned !== nextProps.item.isPinned) return false;
  if (prevProps.item.linkUrl !== nextProps.item.linkUrl) return false;
  if (prevProps.item.AlertType !== nextProps.item.AlertType) return false;

  if (prevProps.isCarousel !== nextProps.isCarousel) return false;
  if (prevProps.currentIndex !== nextProps.currentIndex) return false;
  if (prevProps.totalAlerts !== nextProps.totalAlerts) return false;

  if (prevProps.alertType !== nextProps.alertType) return false;

  if (prevProps.userTargetingEnabled !== nextProps.userTargetingEnabled)
    return false;

  if (prevProps.remove !== nextProps.remove) return false;
  if (prevProps.hideForever !== nextProps.hideForever) return false;
  if (prevProps.onNext !== nextProps.onNext) return false;
  if (prevProps.onPrevious !== nextProps.onPrevious) return false;

  return true;
};

export default React.memo(AlertItem, arePropsEqual);
