import * as React from "react";
import { logger } from "../Services/LoggerService";
import { MessageBar, MessageBarType } from "@fluentui/react";
import styles from "./Alerts.module.scss";
import { IAlertsProps, IAlertType, AlertPriority } from "./IAlerts";
import AlertItem from "../AlertItem/AlertItem";
import AlertListShimmer from "../UI/AlertListShimmer";
import AlertSettingsTabs from "../Settings/AlertSettingsTabs";
import { ISettingsData } from "../Settings/Tabs/SettingsTab";
import { EditModeDetector } from "../Utils/EditModeDetector";
import { useAlertsState, useAlertsDispatch } from "../Context/AlertsContext";
import { StorageService } from "../Services/StorageService";
import { ArrayUtils } from "../Utils/ArrayUtils";
import { StringUtils } from "../Utils/StringUtils";
import { CAROUSEL_CONFIG } from "../Utils/AppConstants";
import { ErrorBoundary } from "../Utils/ErrorBoundary";
import * as strings from "AlertBannerApplicationCustomizerStrings";

const Alerts: React.FC<IAlertsProps> = (props) => {
  const state = useAlertsState();
  const {
    initializeAlerts,
    removeAlert,
    hideAlertForever,
    updateCarouselSettings,
  } = useAlertsDispatch();
  const { alerts, alertTypes, isLoading, hasError, errorMessage } = state;

  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [isInEditMode, setIsInEditMode] = React.useState(false);

  const { carouselEnabled, carouselInterval } = state;
  const carouselTimer = React.useRef<number | null>(null);
  const storageService = React.useRef<StorageService>(
    StorageService.getInstance(),
  );

  const defaultAlertType = React.useMemo<IAlertType>(
    () => ({
      name: strings.DefaultAlertTypeName,
      iconName: "Info",
      backgroundColor: "#ffffff",
      textColor: "#323130",
      additionalStyles: "",
      defaultPriority: AlertPriority.Medium,
      priorityStyles: {
        [AlertPriority.Critical]: "border: 4px solid #a4262c;",
        [AlertPriority.High]: "border: 3px solid #ca5010;",
        [AlertPriority.Medium]: "border: 2px solid #0078d4;",
        [AlertPriority.Low]: "border: 1px solid #107c10;",
      },
    }),
    [],
  );

  const normalizedSiteIds = React.useMemo(() => {
    return (props.siteIds ?? [])
      .map((id) => (id ?? "").toString().trim())
      .filter((id) => StringUtils.isNotEmpty(id));
  }, [props.siteIds]);

  const uniqueSortedSiteIds = React.useMemo(() => {
    return ArrayUtils.unique(normalizedSiteIds).sort();
  }, [normalizedSiteIds]);

  const previousInitPropsRef = React.useRef<{
    siteIds: string[];
    alertTypesJson: string;
    userTargetingEnabled: boolean;
    notificationsEnabled: boolean;
    graphClient: typeof props.graphClient;
    context: typeof props.context;
  } | null>(null);

  React.useEffect(() => {
    const nextInitProps = {
      siteIds: uniqueSortedSiteIds,
      alertTypesJson: props.alertTypesJson,
      userTargetingEnabled: !!props.userTargetingEnabled,
      notificationsEnabled: !!props.notificationsEnabled,
      graphClient: props.graphClient,
      context: props.context,
    };

    const prevInitProps = previousInitPropsRef.current;
    const hasInitPropsChanged =
      !prevInitProps ||
      prevInitProps.graphClient !== nextInitProps.graphClient ||
      prevInitProps.context !== nextInitProps.context ||
      prevInitProps.alertTypesJson !== nextInitProps.alertTypesJson ||
      prevInitProps.userTargetingEnabled !==
        nextInitProps.userTargetingEnabled ||
      prevInitProps.notificationsEnabled !==
        nextInitProps.notificationsEnabled ||
      prevInitProps.siteIds.length !== nextInitProps.siteIds.length ||
      prevInitProps.siteIds.some(
        (id, index) => id !== nextInitProps.siteIds[index],
      );

    if (hasInitPropsChanged) {
      previousInitPropsRef.current = {
        ...nextInitProps,
        siteIds: [...nextInitProps.siteIds],
      };

      initializeAlerts({
        graphClient: props.graphClient,
        context: props.context,
        siteIds: uniqueSortedSiteIds,
        alertTypesJson: props.alertTypesJson,
        userTargetingEnabled: !!props.userTargetingEnabled,
        notificationsEnabled: !!props.notificationsEnabled,
        enableTargetSite: !!props.enableTargetSite,
      });
    }

    setIsInEditMode(EditModeDetector.isPageInEditMode(props.context));
  }, [
    props.graphClient,
    props.context,
    props.alertTypesJson,
    props.userTargetingEnabled,
    props.notificationsEnabled,
    props.enableTargetSite,
    uniqueSortedSiteIds,
    initializeAlerts,
  ]);

  React.useEffect(() => {
    let debounceTimer: number | null = null;
    let intervalTimer: number | null = null;

    const checkEditMode = () => {
      const newEditMode = EditModeDetector.isPageInEditMode(props.context);
      setIsInEditMode((prevMode) => {
        if (prevMode !== newEditMode) {
          logger.debug(
            "Alerts",
            `Edit mode changed: ${prevMode} -> ${newEditMode}`,
          );
        }
        return newEditMode;
      });
    };

    const debouncedCheckEditMode = () => {
      if (debounceTimer) {
        window.clearTimeout(debounceTimer);
      }
      debounceTimer = window.setTimeout(checkEditMode, 50);
    };

    const observer = new MutationObserver(debouncedCheckEditMode);

    observer.observe(document.body, {
      attributes: true,
      attributeFilter: ["class"],
      subtree: false,
    });

    const observeCommandBar = () => {
      const commandBar = document.querySelector(
        '[data-automation-id="pageCommandBarRegion"]',
      );
      if (commandBar) {
        observer.observe(commandBar, {
          childList: true,
          subtree: true,
        });
      }
    };

    observeCommandBar();

    const retryTimer = window.setTimeout(observeCommandBar, 1000);

    intervalTimer = window.setInterval(checkEditMode, 500);

    checkEditMode();
    return () => {
      if (debounceTimer) {
        window.clearTimeout(debounceTimer);
      }
      if (intervalTimer) {
        window.clearInterval(intervalTimer);
      }
      window.clearTimeout(retryTimer);
      observer.disconnect();
    };
  }, [props.context]);

  // Effect to reset index when alerts count changes (not on every array reference change)
  const prevAlertsLengthRef = React.useRef(alerts.length);
  React.useEffect(() => {
    if (alerts.length !== prevAlertsLengthRef.current) {
      prevAlertsLengthRef.current = alerts.length;
      if (alerts.length > 0 && currentIndex >= alerts.length) {
        setCurrentIndex(alerts.length - 1);
      } else if (alerts.length === 0) {
        setCurrentIndex(0);
      }
    }
  }, [alerts.length, currentIndex]);

  const alertsLengthRef = React.useRef(alerts.length);
  React.useEffect(() => {
    alertsLengthRef.current = alerts.length;
  }, [alerts.length]);

  React.useEffect(() => {
    if (carouselEnabled && alerts.length > 1) {
      carouselTimer.current = window.setInterval(() => {
        setCurrentIndex(
          (prevIndex) => (prevIndex + 1) % alertsLengthRef.current,
        );
      }, carouselInterval);
    } else if (carouselTimer.current) {
      window.clearInterval(carouselTimer.current);
      carouselTimer.current = null;
    }

    return () => {
      if (carouselTimer.current) {
        window.clearInterval(carouselTimer.current);
        carouselTimer.current = null;
      }
    };
  }, [carouselEnabled, carouselInterval, alerts.length]);

  React.useEffect(() => {
    const savedCarouselEnabled =
      storageService.current.getFromLocalStorage<boolean>("carouselEnabled");
    const savedCarouselInterval =
      storageService.current.getFromLocalStorage<number>("carouselInterval");

    const settings: { carouselEnabled?: boolean; carouselInterval?: number } =
      {};

    if (savedCarouselEnabled !== null) {
      settings.carouselEnabled = savedCarouselEnabled;
    }
    if (
      savedCarouselInterval &&
      savedCarouselInterval >= CAROUSEL_CONFIG.MIN_INTERVAL &&
      savedCarouselInterval <= CAROUSEL_CONFIG.MAX_INTERVAL
    ) {
      settings.carouselInterval = savedCarouselInterval;
    }

    if (Object.keys(settings).length > 0) {
      updateCarouselSettings(settings);
    }
  }, [updateCarouselSettings]);

  const handleSettingsChange = React.useCallback(
    (settings: ISettingsData & { enableTargetSite?: boolean }) => {
      if (props.onSettingsChange) {
        props.onSettingsChange(settings);
      }
      // The context will handle reloading alert types if they changed via its own logic
    },
    [props.onSettingsChange],
  );

  const goToNext = React.useCallback(() => {
    setCurrentIndex((prevIndex) => (prevIndex + 1) % alerts.length);
  }, [alerts.length]);

  const goToPrevious = React.useCallback(() => {
    setCurrentIndex(
      (prevIndex) => (prevIndex - 1 + alerts.length) % alerts.length,
    );
  }, [alerts.length]);

  const handleMouseEnter = React.useCallback(() => {
    if (carouselTimer.current) {
      window.clearInterval(carouselTimer.current);
      carouselTimer.current = null;
    }
  }, []);

  const handleMouseLeave = React.useCallback(() => {
    if (carouselEnabled && alertsLengthRef.current > 1) {
      carouselTimer.current = window.setInterval(() => {
        setCurrentIndex(
          (prevIndex) => (prevIndex + 1) % alertsLengthRef.current,
        );
      }, carouselInterval);
    }
  }, [carouselEnabled, carouselInterval]);

  const hasAlerts = alerts.length > 0;

  return (
    <div
      className={`${styles.alerts} ${isLoading ? (styles as any).loading : ""}`}
      aria-live="polite"
      aria-atomic="true"
      role="region"
      aria-label="Organization Alerts"
    >
      {isLoading && !hasAlerts && !hasError && (
        <div style={{ padding: "8px 16px" }}>
          <AlertListShimmer rowCount={1} />
        </div>
      )}
      {hasError && (
        <div className={`${styles.errorContainer} ${styles.errorWrapper}`}>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={true}
            className={styles.errorMessageBar}
          >
            <div className={styles.errorTitle}>
              {strings.AlertsLoadErrorTitle}
            </div>
            <div className={styles.errorDetail}>
              {errorMessage || strings.AlertsLoadErrorFallback}
            </div>
          </MessageBar>
        </div>
      )}
      {hasAlerts && (
        <div
          className={styles.carousel}
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
        >
          <ErrorBoundary
            componentName="AlertItem"
            onError={(error) =>
              logger.error("Alerts", "Alert rendering failed", error)
            }
          >
            <AlertItem
              key={alerts[currentIndex].id}
              item={alerts[currentIndex]}
              remove={removeAlert}
              hideForever={hideAlertForever}
              alertType={
                alertTypes[alerts[currentIndex].AlertType] || defaultAlertType
              }
              isCarousel={true}
              currentIndex={currentIndex + 1}
              totalAlerts={alerts.length}
              onNext={goToNext}
              onPrevious={goToPrevious}
              userTargetingEnabled={props.userTargetingEnabled || false}
            />
          </ErrorBoundary>
        </div>
      )}
      {isInEditMode && (
        <AlertSettingsTabs
          isInEditMode={isInEditMode}
          alertTypesJson={props.alertTypesJson}
          userTargetingEnabled={props.userTargetingEnabled || false}
          notificationsEnabled={props.notificationsEnabled || false}
          enableTargetSite={props.enableTargetSite || false}
          emailServiceAccount={props.emailServiceAccount}
          copilotEnabled={props.copilotEnabled || false}
          graphClient={props.graphClient}
          context={props.context}
          onSettingsChange={handleSettingsChange}
        />
      )}
    </div>
  );
};

export default Alerts;
