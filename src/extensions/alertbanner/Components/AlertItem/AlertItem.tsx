import * as React from "react";
import {
  IAlertItem,
  IAlertType,
  IQuickAction
} from "../Alerts/IAlerts";
import styles from "./AlertItem.module.scss";

// Import new components
import AlertHeader from "./AlertHeader";
import AlertContent from "./AlertContent";
import AlertActions from "./AlertActions";
import { PRIORITY_COLORS } from "./utils"; // Import from utils.tsx



// ===== UTILITY FUNCTIONS =====


const parseAdditionalStyles = (stylesString?: string): React.CSSProperties => {
  if (!stylesString) return {};
  
  const styleObj: Record<string, string | number> = {};
  const stylesArray = stylesString.split(";").filter(s => s.trim());
  
  stylesArray.forEach(style => {
    const [key, value] = style.split(":");
    if (key?.trim() && value?.trim()) {
      const camelCaseKey = key.trim().replace(/-([a-z])/g, (_, group1) => group1.toUpperCase());
      const trimmedValue = value.trim();
      styleObj[camelCaseKey] = isNaN(Number(trimmedValue)) ? trimmedValue : Number(trimmedValue);
    }
  });
  
  return styleObj as React.CSSProperties;
};



// ===== INTERFACES =====
export interface IAlertItemProps {
  item: IAlertItem;
  remove: (id: number) => void;
  hideForever: (id: number) => void;
  alertType: IAlertType;
  richMediaEnabled?: boolean;
  // Carousel props
  isCarousel?: boolean;
  currentIndex?: number;
  totalAlerts?: number;
  onNext?: () => void;
  onPrevious?: () => void;
}






const AlertItem: React.FC<IAlertItemProps> = ({
  item,
  remove,
  hideForever,
  alertType,
  richMediaEnabled = false,
  isCarousel = false,
  currentIndex = 1,
  totalAlerts = 1,
  onNext,
  onPrevious
}) => {
  // Component state
  const [expanded, setExpanded] = React.useState(false);

  // Accessibility IDs
  const ariaControlsId = `alert-description-${item.Id}`;


  // Event handlers
  const handlers = React.useMemo(() => ({
    toggleExpanded: () => setExpanded(prev => !prev),
    remove: (id: number) => remove(id),
    hideForever: (id: number) => hideForever(id),
    stopPropagation: (e: React.MouseEvent) => e.stopPropagation(),
  }), [remove, hideForever]);

  const handleQuickAction = React.useCallback((action: IQuickAction) => {
    switch (action.actionType) {
      case "link":
        if (action.url) {
          window.open(action.url, "_blank", "noopener,noreferrer");
        }
        break;
      case "dismiss":
        handlers.remove(item.Id); // Use handlers.remove with item.Id
        break;
      case "acknowledge":
        console.log(`Alert ${item.Id} acknowledged`);
        handlers.remove(item.Id); // Use handlers.remove with item.Id
        break;
      case "custom":
        if (action.callback && typeof (window as any)[action.callback] === "function") {
          (window as any)[action.callback](item);
        }
        break;
    }
  }, [handlers.remove, item]);

  


  const baseContainerStyle = React.useMemo<React.CSSProperties>(() => ({
    backgroundColor: alertType.backgroundColor || "#389899",
    color: alertType.textColor || "#ffffff",
    ...parseAdditionalStyles(alertType.additionalStyles)
  }), [alertType]);

  const priorityStyle = React.useMemo(
    () =>
      alertType.priorityStyles
        ? alertType.priorityStyles[item.priority as keyof typeof alertType.priorityStyles]
        : "",
    [alertType.priorityStyles, item.priority]
  );

  const containerStyle = React.useMemo<React.CSSProperties>(() => ({
    ...baseContainerStyle,
    background: `linear-gradient(to right, ${PRIORITY_COLORS[item.priority].start}, ${PRIORITY_COLORS[item.priority].end})`,
    ...parseAdditionalStyles(priorityStyle),
    ...(item.priority === "critical" && {
      border: '2px solid #E81123',
      boxShadow: '0 0 10px rgba(232, 17, 35, 0.5)'
    })
  }), [baseContainerStyle, priorityStyle, item.priority]);

  const containerClassNames = [
    styles.container,
    styles.clickable,
    item.priority === "critical" ? styles.critical : '',
    item.priority === "high" ? styles.high : '',
    item.priority === "medium" ? styles.medium : '',
    item.priority === "low" ? styles.low : '',
    item.isPinned ? styles.pinned : ''
  ].filter(Boolean).join(' ');

  

  // Use native Fluent UI v9 dialog styling - no custom overrides needed

  return (
    <div className={styles.alertItem}>
      <div
        className={containerClassNames}
        style={containerStyle}
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
          expanded={expanded}
          toggleExpanded={handlers.toggleExpanded}
          ariaControlsId={ariaControlsId}
        />
        <AlertContent
          item={item}
          richMediaEnabled={richMediaEnabled}
          expanded={expanded}
          stopPropagation={handlers.stopPropagation}
          handleQuickAction={handleQuickAction}
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
        />
      </div>
    </div>
  );
};

export default AlertItem;