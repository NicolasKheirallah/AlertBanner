import * as React from "react";
import { MessageBar, MessageBarBody, MessageBarTitle, tokens } from "@fluentui/react-components";
import styles from "./Alerts.module.scss";
import { IAlertsProps, IAlertType, AlertPriority } from "./IAlerts";
import AlertItem from "../AlertItem/AlertItem";
import AlertSettings, { ISettingsData } from "../Settings/AlertSettings";
import { EditModeDetector } from "../Utils/EditModeDetector";
import { useAlerts } from "../Context/AlertsContext";

const Alerts: React.FC<IAlertsProps> = (props) => {
  const { state, initializeAlerts, removeAlert, hideAlertForever } = useAlerts();
  const { alerts, alertTypes, isLoading, hasError, errorMessage } = state;

  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [isInEditMode, setIsInEditMode] = React.useState(false);

  // Initialize alerts and edit mode detection on mount
  React.useEffect(() => {
    initializeAlerts({
      graphClient: props.graphClient,
      siteIds: props.siteIds,
      alertTypesJson: props.alertTypesJson,
      userTargetingEnabled: props.userTargetingEnabled,
      notificationsEnabled: props.notificationsEnabled,
      richMediaEnabled: props.richMediaEnabled,
    });

    setIsInEditMode(EditModeDetector.isPageInEditMode());
    const cleanup = EditModeDetector.onEditModeChange(setIsInEditMode);
    return cleanup;
  }, [props.graphClient, props.siteIds, props.alertTypesJson, props.userTargetingEnabled, props.notificationsEnabled, props.richMediaEnabled, initializeAlerts]);

  // Effect to reset index when alerts change
  React.useEffect(() => {
    if (alerts.length > 0 && currentIndex >= alerts.length) {
      setCurrentIndex(alerts.length - 1);
    } else if (alerts.length === 0) {
      setCurrentIndex(0);
    }
  }, [alerts, currentIndex]);

  const handleSettingsChange = (settings: ISettingsData) => {
    if (props.onSettingsChange) {
      props.onSettingsChange(settings);
    }
    // The context will handle reloading alert types if they changed via its own logic
  };

  // Carousel navigation
  const goToNext = () => {
    setCurrentIndex((prevIndex) => (prevIndex + 1) % alerts.length);
  };

  const goToPrevious = () => {
    setCurrentIndex((prevIndex) => (prevIndex - 1 + alerts.length) % alerts.length);
  };

  if (isLoading) {
    return null; // Hide loading, let alerts load silently in the background
  }

  if (hasError) {
    return (
      <div style={{
        width: '100%',
        maxWidth: '1200px',
        margin: '0 auto',
        padding: tokens.spacingVerticalM,
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalM,
      }}>
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Unable to Load Alerts</MessageBarTitle>
            {errorMessage || "An error occurred while loading alerts. Please try refreshing the page."}
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  const hasAlerts = alerts.length > 0;

  if (!hasAlerts && !isInEditMode) {
    return null; // Hide component completely if no alerts and not in edit mode
  }

  return (
    <div className={styles.alerts}>
      {hasAlerts && (
        <div className={styles.carousel}>
          <AlertItem
            key={alerts[currentIndex].id}
            item={alerts[currentIndex]}
            remove={removeAlert}
            hideForever={hideAlertForever}
            alertType={alertTypes[alerts[currentIndex].AlertType] || defaultAlertType}
            richMediaEnabled={props.richMediaEnabled}
            isCarousel={true}
            currentIndex={currentIndex + 1}
            totalAlerts={alerts.length}
            onNext={goToNext}
            onPrevious={goToPrevious}
          />
        </div>
      )}
      {isInEditMode && (
        <AlertSettings
          isInEditMode={isInEditMode}
          alertTypesJson={props.alertTypesJson}
          userTargetingEnabled={props.userTargetingEnabled || false}
          notificationsEnabled={props.notificationsEnabled || false}
          richMediaEnabled={props.richMediaEnabled || false}
          graphClient={props.graphClient}
          context={props.context}
          onSettingsChange={handleSettingsChange}
        />
      )}
    </div>
  );
};

// Define a default alert type in case an alert type is missing
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
    [AlertPriority.Low]: "",
  },
};

export default Alerts;
