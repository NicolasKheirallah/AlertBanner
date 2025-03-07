import * as React from "react";
import { useEffect, memo } from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { IAlertsProps, IAlertType, AlertPriority } from "./IAlerts";
import AlertItem from "../AlertItem/AlertItem";
import { useAlerts } from "../Context/AlertsContext";
import styles from "./Alerts.module.scss";

// Create a memoized version of AlertItem to prevent unnecessary re-renders
const MemoizedAlertItem = memo(AlertItem, (prevProps, nextProps) => {
  // Only re-render if these props change
  return (
    prevProps.item.Id === nextProps.item.Id &&
    prevProps.item.title === nextProps.item.title &&
    prevProps.item.description === nextProps.item.description &&
    prevProps.item.isPinned === nextProps.item.isPinned &&
    prevProps.alertType === nextProps.alertType
  );
});

// Define an error boundary component
class ErrorBoundary extends React.Component<{children: React.ReactNode, fallback?: React.ReactNode}, {hasError: boolean}> {
  constructor(props: {children: React.ReactNode, fallback?: React.ReactNode}) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(_: Error) {
    return { hasError: true };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("Alert Banner Error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return this.props.fallback || (
        <div style={{padding: '10px', color: '#666', fontSize: '13px'}}>
          Unable to load alerts. Please try refreshing the page.
        </div>
      );
    }

    return this.props.children;
  }
}

/**
 * AlertsWithContext - Component that uses Context API but with performance optimizations
 */
const AlertsWithContext: React.FC<IAlertsProps> = (props) => {
  const { 
    state, 
    removeAlert, 
    hideAlertForever, 
    initializeAlerts,
    refreshAlerts
  } = useAlerts();
  
  const { 
    alerts, 
    alertTypes, 
    isLoading, 
    hasError, 
    errorMessage 
  } = state;

  // Initialize on component mount
  useEffect(() => {
    try {
      initializeAlerts({
        graphClient: props.graphClient,
        siteIds: props.siteIds || [],
        alertTypesJson: props.alertTypesJson,
        userTargetingEnabled: props.userTargetingEnabled,
        notificationsEnabled: props.notificationsEnabled,
        richMediaEnabled: props.richMediaEnabled
      });
    } catch (error) {
      console.error("Error initializing alerts:", error);
    }

    // Set up automatic refresh interval
    const refreshInterval = setInterval(() => {
      try {
        refreshAlerts();
      } catch (error) {
        console.error("Error refreshing alerts:", error);
      }
    }, 5 * 60 * 1000); // 5 minutes

    // Clean up on unmount
    return () => {
      clearInterval(refreshInterval);
    };
  }, [
    props.graphClient,
    props.siteIds,
    props.alertTypesJson,
    props.userTargetingEnabled,
    props.notificationsEnabled,
    props.richMediaEnabled,
    initializeAlerts,
    refreshAlerts
  ]);

  // Default alert type
  const defaultAlertType: IAlertType = {
    name: "Default",
    iconName: "Info",
    backgroundColor: "#ffffff",
    textColor: "#000000",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  };

  // Error boundary fallback
  const errorFallback = (
    <div className={styles.alerts}>
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={false}
        dismissButtonAriaLabel="Close"
      >
        An error occurred while loading alerts. Please try refreshing the page.
      </MessageBar>
    </div>
  );

  // Render loading spinner
  if (isLoading) {
    return (
      <div className={styles.alerts}>
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.medium} label="Loading alerts..." />
        </div>
      </div>
    );
  }

  // Render error message
  if (hasError) {
    return (
      <div className={styles.alerts}>
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
        >
          {errorMessage || "An error occurred while loading alerts."}
        </MessageBar>
      </div>
    );
  }

  // Check if we have alerts to show
  const hasAlerts = alerts.length > 0;

  // Render alerts with error boundary
  return (
    <ErrorBoundary fallback={errorFallback}>
      <div className={styles.alerts}>
        {hasAlerts ? (
          <div className={styles.container}>
            {alerts.map((alert) => {
              const alertType = alertTypes[alert.AlertType] || defaultAlertType;
              return (
                <MemoizedAlertItem
                  key={alert.Id}
                  item={alert}
                  remove={removeAlert}
                  hideForever={hideAlertForever}
                  alertType={alertType}
                  richMediaEnabled={props.richMediaEnabled}
                />
              );
            })}
          </div>
        ) : (
          <div className={styles.noAlerts}>
            <MessageBar messageBarType={MessageBarType.info}>
              No alerts to display.
            </MessageBar>
          </div>
        )}
      </div>
    </ErrorBoundary>
  );
};

export default AlertsWithContext;