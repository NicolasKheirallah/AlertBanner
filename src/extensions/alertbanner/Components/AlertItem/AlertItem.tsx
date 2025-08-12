import * as React from "react";
import {
  Button,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  Dismiss24Regular,
  ChevronDown24Regular,
  ChevronUp24Regular,
  EyeOff24Regular,
  Link24Regular,
  Info24Regular,
  Warning24Regular,
  ErrorCircle24Regular,
  CheckmarkCircle24Regular,
  ChevronLeft24Regular,
  ChevronRight24Regular
} from "@fluentui/react-icons";
import {
  IAlertItem,
  IAlertType,
  AlertPriority,
  IQuickAction
} from "../Alerts/IAlerts";
import RichMediaAlert from "../Services/RichMediaAlert";
import styles from "./AlertItem.module.scss";
import richMediaStyles from "../Services/RichMediaAlert.module.scss";

// ===== CONSTANTS =====
const PRIORITY_COLORS = {
  critical: '#d13438',
  high: '#f7630c',
  medium: '#0078d4',
  low: '#107c10'
} as const;





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

const getPriorityIcon = (priority: AlertPriority): React.ReactElement => {
  const iconStyle = { width: 20, height: 20, color: PRIORITY_COLORS[priority] || PRIORITY_COLORS.low };
  
  switch (priority) {
    case "critical":
      return <ErrorCircle24Regular style={iconStyle} />;
    case "high":
      return <Warning24Regular style={iconStyle} />;
    case "medium":
      return <Info24Regular style={iconStyle} />;
    case "low":
    default:
      return <CheckmarkCircle24Regular style={iconStyle} />;
  }
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


// Sub-components
interface IDescriptionContentProps {
  description: string;
}

const DescriptionContent: React.FC<IDescriptionContentProps> = React.memo(({ description }) => {
  const [isExpanded, setIsExpanded] = React.useState(false);
  const TRUNCATE_LENGTH = 200; // Character limit for truncation

  const toggleExpanded = () => {
    setIsExpanded(!isExpanded);
  };

  let displayedDescription = description;
  let showReadMoreButton = false;

  // Only truncate if it's not HTML and it's longer than the limit
  if (!/<[a-z][\s\S]*>/i.test(description) && description.length > TRUNCATE_LENGTH && !isExpanded) {
    displayedDescription = description.substring(0, TRUNCATE_LENGTH) + "...";
    showReadMoreButton = true;
  }

  // If description contains HTML tags, render it directly
  if (/<[a-z][\s\S]*>/i.test(description)) {
    return (
      <div
        className={richMediaStyles.markdownContainer}
        dangerouslySetInnerHTML={{ __html: description }}
      />
    );
  }

  const paragraphs = displayedDescription.split("\n\n");

  return (
    <div className={richMediaStyles.markdownContainer}>
      {paragraphs.map((paragraph, index) => {
        // Handle lists
        if (paragraph.includes("\n- ") || paragraph.includes("\n* ")) {
          const [listTitle, ...listItems] = paragraph.split(/\n[-*]\s+/);
          return (
            <div key={`para-${index}`} style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
              {listTitle.trim() && <Text>{listTitle.trim()}</Text>}
              {listItems.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
                  {listItems.map((listItem, itemIndex) => (
                    <div
                      key={`list-item-${itemIndex}`}
                      style={{ display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'flex-start' }}
                    >
                      <Text>•</Text>
                      <Text>{listItem.trim()}</Text>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        }

        // Handle bold text
        if (paragraph.includes("**") || paragraph.includes("__")) {
          const parts = paragraph.split(/(\**.*?\**|__.*?__)/g);
          return (
            <Text key={`para-${index}`}>
              {parts.map((part, partIndex) => {
                const isBold = (part.startsWith("**") && part.endsWith("**")) ||
                              (part.startsWith("__") && part.endsWith("__"));
                
                if (isBold) {
                  return (
                    <span
                      key={`part-${partIndex}`}
                      style={{ fontWeight: tokens.fontWeightSemibold }}
                    >
                      {part.slice(2, -2)}
                    </span>
                  );
                }
                return part;
              })}
            </Text>
          );
        }

        // Simple paragraph
        return <Text key={`para-${index}`}>{paragraph}</Text>;
      })}
      {(showReadMoreButton || (description.length > TRUNCATE_LENGTH && isExpanded)) && (
        <Button
          appearance="transparent"
          size="small"
          onClick={toggleExpanded}
          style={{ alignSelf: 'flex-start', marginTop: tokens.spacingVerticalS }}
        >
          {isExpanded ? "Show Less" : "Read More"}
        </Button>
      )}
    </div>
  );
});

interface IAlertLinkProps {
  link?: {
    Url: string;
    Description: string;
  };
  isDialog?: boolean;
  onClick?: (e: React.MouseEvent) => void;
}

const AlertLinkButton: React.FC<IAlertLinkProps> = React.memo(({ link, isDialog = false, onClick }) => {
  if (!link) return null;

  const buttonStyle: React.CSSProperties = isDialog 
    ? {
        backgroundColor: '#0078d4',
        border: '1px solid #0078d4',
        color: '#ffffff',
        borderRadius: '2px',
        padding: '6px 16px',
        fontSize: '14px'
      }
    : {
        padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
        backgroundColor: 'rgba(255, 255, 255, 0.1)',
        borderRadius: tokens.borderRadiusMedium,
      };

  return (
    <Button
      icon={<Link24Regular />}
      as="a"
      href={link.Url}
      target="_blank"
      rel="noopener noreferrer"
      onClick={onClick}
      appearance={isDialog ? "primary" : "secondary"}
      style={buttonStyle}
    >
      {link.Description}
    </Button>
  );
});

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
    remove: () => remove(item.Id),
    hideForever: () => hideForever(item.Id),
    stopPropagation: (e: React.MouseEvent) => e.stopPropagation(),
  }), [item.Id, remove, hideForever]);

  const handleQuickAction = React.useCallback((action: IQuickAction) => {
    switch (action.actionType) {
      case "link":
        if (action.url) {
          window.open(action.url, "_blank", "noopener,noreferrer");
        }
        break;
      case "dismiss":
        handlers.remove();
        break;
      case "acknowledge":
        console.log(`Alert ${item.Id} acknowledged`);
        handlers.remove();
        break;
      case "custom":
        if (action.callback && typeof (window as any)[action.callback] === "function") {
          (window as any)[action.callback](item);
        }
        break;
    }
  }, [handlers.remove, item]);

  const renderQuickActions = React.useCallback(() => {
    if (!item.quickActions?.length) return null;
    return (
      <div
        style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalS }}
        onClick={handlers.stopPropagation}
      >
        {item.quickActions.map((action, index) => {
          return (
            <Button
              key={`${item.Id}-action-${index}`}
              appearance="outline"
              size="small"
              icon={<Link24Regular />}
              onClick={(e) => {
                handlers.stopPropagation(e);
                handleQuickAction(action);
              }}
                  >
              {action.label}
            </Button>
          );
        })}
      </div>
    );
  }, [item.quickActions, handlers.stopPropagation, handleQuickAction, item.Id]);


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

  const descriptionClassName = expanded ? styles.alertDescriptionExp : styles.alertDescription;

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
          {item.description && (
            <div className={descriptionClassName} id={ariaControlsId}>
              {expanded ? (
                <DescriptionContent description={item.description} />
              ) : (
                <div
                  className={styles.truncatedHtml}
                  dangerouslySetInnerHTML={{ __html: item.description }}
                />
              )}
            </div>
          )}
          {item.richMedia && richMediaEnabled && (
            <div onClick={handlers.stopPropagation}>
              <RichMediaAlert media={item.richMedia} expanded={expanded} />
            </div>
          )}
          {expanded && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
              {item.richMedia && richMediaEnabled && (
                <div style={{ marginTop: '16px' }}>
                  <RichMediaAlert media={item.richMedia} expanded={true} />
                </div>
              )}
              
              {item.link && (
                <div onClick={handlers.stopPropagation}>
                  <AlertLinkButton 
                    link={item.link} 
                    onClick={handlers.stopPropagation}
                  />
                </div>
              )}
              {renderQuickActions()}
            </div>
          )}
        </div>
        <div className={styles.actionSection} onClick={handlers.stopPropagation}>
          {isCarousel && totalAlerts > 1 && (
            <>
              <Button
                appearance="transparent"
                icon={<ChevronLeft24Regular />}
                onClick={onPrevious}
                aria-label="Previous Alert"
                size="small"
              />
              <div className={styles.carouselCounter}>
                <Text size={200} weight="medium" style={{ color: tokens.colorNeutralForeground2 }}>
                  {currentIndex} of {totalAlerts}
                </Text>
              </div>
              <Button
                appearance="transparent"
                icon={<ChevronRight24Regular />}
                onClick={onNext}
                aria-label="Next Alert"
                size="small"
              />
              <div className={styles.divider} />
            </>
          )}
          <Button
            appearance="transparent"
            icon={expanded ? <ChevronUp24Regular /> : <ChevronDown24Regular />}
            onClick={handlers.toggleExpanded}
            aria-expanded={expanded}
            aria-controls={ariaControlsId}
            aria-label={expanded ? "Collapse Alert" : "Expand Alert"}
            size="small"
          />
          <Button
            appearance="transparent"
            icon={<Dismiss24Regular />}
            onClick={handlers.remove}
            aria-label="Dismiss Alert"
            size="small"
          />
          <Button
            appearance="transparent"
            icon={<EyeOff24Regular />}
            onClick={handlers.hideForever}
            aria-label="Hide Alert Forever"
            size="small"
          />
        </div>
      </div>
    </div>
  );
};

export default AlertItem;