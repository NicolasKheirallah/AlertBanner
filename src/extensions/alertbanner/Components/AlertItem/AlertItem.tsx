import * as React from "react";
import {
  Button,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
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

// Utility functions
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
  const iconProps = { style: { width: 20, height: 20 } };
  
  switch (priority) {
    case "critical":
      return <ErrorCircle24Regular {...iconProps} style={{ ...iconProps.style, color: '#d13438' }} />;
    case "high":
      return <Warning24Regular {...iconProps} style={{ ...iconProps.style, color: '#f7630c' }} />;
    case "medium":
      return <Info24Regular {...iconProps} style={{ ...iconProps.style, color: '#0078d4' }} />;
    case "low":
    default:
      return <CheckmarkCircle24Regular {...iconProps} style={{ ...iconProps.style, color: '#107c10' }} />;
  }
};

const getPriorityTitleColor = (priority: AlertPriority): string => {
  switch (priority) {
    case "critical":
      return '#d13438';
    case "high":
      return '#f7630c';
    case "medium":
      return '#0078d4';
    case "low":
    default:
      return '#107c10';
  }
};

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
  // If description contains HTML tags, render it directly
  if (/<[a-z][\s\S]*>/i.test(description)) {
    return (
      <div
        className={richMediaStyles.markdownContainer}
        dangerouslySetInnerHTML={{ __html: description }}
      />
    );
  }

  const paragraphs = description.split("\n\n");
  const containerStyle = {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalM
  };

  return (
    <div className={richMediaStyles.markdownContainer} style={containerStyle}>
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
          const parts = paragraph.split(/(\*\*.*?\*\*|__.*?__)/g);
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
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);

  // Accessibility IDs
  const ariaControlsId = `alert-description-${item.Id}`;
  const dialogTitleId = `dialog-title-${item.Id}`;
  const dialogContentId = `dialog-content-${item.Id}`;




  // Event handlers
  const handlers = React.useMemo(() => ({
    toggleExpanded: () => setExpanded(prev => !prev),
    openDialog: () => setIsDialogOpen(true),
    closeDialog: () => setIsDialogOpen(false),
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
        onClick={handlers.openDialog}
        role="button"
        tabIndex={0}
        aria-expanded={isDialogOpen}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            handlers.openDialog();
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
      <Dialog 
        open={isDialogOpen} 
        onOpenChange={(event, data) => !data.open && handlers.closeDialog()}
        modalType="modal"
      >
        <DialogSurface 
          style={{
            maxWidth: '630px',
            minWidth: '340px',
            backgroundColor: '#ffffff',
            border: 'none',
            borderRadius: '2px',
            boxShadow: '0 6.4px 14.4px 0 rgba(0, 0, 0, 0.132), 0 1.2px 3.6px 0 rgba(0, 0, 0, 0.108)',
            outline: 'none'
          }}
        >
          <DialogBody style={{ padding: '0', border: 'none', outline: 'none' }}>
            <div 
              style={{
                backgroundColor: '#ffffff',
                borderBottom: '1px solid #edebe9',
                padding: '20px 24px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                minHeight: '60px'
              }}
            >
              <DialogTitle 
                id={dialogTitleId}
                as="div"
                style={{ 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  color: getPriorityTitleColor(item.priority),
                  fontSize: '20px',
                  fontWeight: '600',
                  fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                  margin: '0',
                  padding: '0',
                  border: 'none',
                  outline: 'none',
                  flex: 1,
                  lineHeight: '28px'
                }}
              >
                {getPriorityIcon(item.priority)}
                {item.title}
              </DialogTitle>
              <button
                onClick={handlers.closeDialog}
                aria-label="Close dialog"
                style={{
                  background: 'transparent',
                  border: 'none',
                  outline: 'none',
                  cursor: 'pointer',
                  padding: '8px',
                  margin: '0',
                  borderRadius: '2px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  color: '#605e5c',
                  fontSize: '16px',
                  minWidth: '32px',
                  minHeight: '32px',
                  transition: 'background-color 0.1s ease'
                }}
                onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f3f2f1'}
                onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
              >
                <Dismiss24Regular style={{ width: '16px', height: '16px' }} />
              </button>
            </div>
            
            <DialogContent 
              id={dialogContentId}
              style={{
                padding: '20px 24px',
                backgroundColor: '#ffffff',
                color: '#323130',
                fontSize: '14px',
                lineHeight: '20px',
                maxHeight: '60vh',
                overflow: 'auto',
                border: 'none',
                outline: 'none',
                fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
              }}
            >
              <DescriptionContent description={item.description} />
              
              {item.richMedia && richMediaEnabled && (
                <div style={{ marginTop: '16px' }}>
                  <RichMediaAlert media={item.richMedia} expanded={true} />
                </div>
              )}
              
              {item.link && (
                <div style={{ marginTop: '16px' }}>
                  <AlertLinkButton link={item.link} isDialog={true} />
                </div>
              )}
            </DialogContent>
            
            <DialogActions 
              style={{
                backgroundColor: '#f3f2f1',
                borderTop: '1px solid #edebe9',
                padding: '16px 24px',
                justifyContent: 'flex-end',
                gap: '8px',
                border: 'none',
                outline: 'none'
              }}
            >
              {item.quickActions?.map((action, index) => {
                if (action.actionType === "dismiss") return null;
                return (
                  <Button
                    key={`dialog-action-${index}`}
                    appearance="primary"
                    onClick={() => {
                      handleQuickAction(action);
                      handlers.closeDialog();
                    }}
                    style={{
                      backgroundColor: getPriorityTitleColor(item.priority),
                      borderColor: getPriorityTitleColor(item.priority)
                    }}
                  >
                    {action.label}
                  </Button>
                );
              })}
              <Button 
                appearance="secondary" 
                onClick={handlers.closeDialog}
                style={{
                  backgroundColor: '#ffffff',
                  border: '1px solid #8a8886',
                  color: '#323130'
                }}
              >
                Close
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default AlertItem;