import * as React from "react";
import { logger } from '../Services/LoggerService';
import { MessageBar, MessageBarBody, MessageBarTitle, tokens } from "@fluentui/react-components";
import styles from "./Alerts.module.scss";
import { IAlertsProps, IAlertType, AlertPriority } from "./IAlerts";
import AlertItem from "../AlertItem/AlertItem";
import AlertSettingsTabs from "../Settings/AlertSettingsTabs";
import { ISettingsData } from "../Settings/Tabs/SettingsTab";
import { EditModeDetector } from "../Utils/EditModeDetector";
import { useAlerts } from "../Context/AlertsContext";
import { StorageService } from "../Services/StorageService";
import { ArrayUtils } from "../Utils/ArrayUtils";
import { StringUtils } from "../Utils/StringUtils";
import { CAROUSEL_CONFIG } from "../Utils/AppConstants";
import { ErrorBoundary } from "../Utils/ErrorBoundary";
import * as strings from 'AlertBannerApplicationCustomizerStrings';

const Alerts: React.FC<IAlertsProps> = (props) => {
  const { state, initializeAlerts, removeAlert, hideAlertForever } = useAlerts();
  const { alerts, alertTypes, isLoading, hasError, errorMessage } = state;

  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [isInEditMode, setIsInEditMode] = React.useState(false);
  
  // Carousel settings
  const [carouselEnabled, setCarouselEnabled] = React.useState(false);
  const [carouselInterval, setCarouselInterval] = React.useState<number>(CAROUSEL_CONFIG.DEFAULT_INTERVAL);
  const carouselTimer = React.useRef<number | null>(null);
  const storageService = React.useRef<StorageService>(StorageService.getInstance());

  const defaultAlertType = React.useMemo<IAlertType>(() => ({
    name: strings.DefaultAlertTypeName,
    iconName: "Info",
    backgroundColor: tokens.colorNeutralBackground1,
    textColor: tokens.colorNeutralForeground1,
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: `border: 2px solid ${tokens.colorPaletteRedForeground1};`,
      [AlertPriority.High]: `border: 1px solid ${tokens.colorPaletteDarkOrangeForeground1};`,
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  }), []);

  // Store initial props to prevent unnecessary re-initialization
  const previousInitPropsRef = React.useRef<{
    siteIds: string[];
    alertTypesJson: string;
    userTargetingEnabled: boolean;
    notificationsEnabled: boolean;
    graphClient: typeof props.graphClient;
    context: typeof props.context;
  } | null>(null);

  // Initialize alerts and edit mode detection on mount
  React.useEffect(() => {
    const normalizedSiteIds = (props.siteIds ?? [])
      .map(id => (id ?? '').toString().trim())
      .filter(id => StringUtils.isNotEmpty(id));
    const uniqueSortedSiteIds = ArrayUtils.unique(normalizedSiteIds).sort();

    const nextInitProps = {
      siteIds: uniqueSortedSiteIds,
      alertTypesJson: props.alertTypesJson,
      userTargetingEnabled: !!props.userTargetingEnabled,
      notificationsEnabled: !!props.notificationsEnabled,
      graphClient: props.graphClient,
      context: props.context
    };

    const prevInitProps = previousInitPropsRef.current;
    const hasInitPropsChanged = !prevInitProps ||
      prevInitProps.graphClient !== nextInitProps.graphClient ||
      prevInitProps.context !== nextInitProps.context ||
      prevInitProps.alertTypesJson !== nextInitProps.alertTypesJson ||
      prevInitProps.userTargetingEnabled !== nextInitProps.userTargetingEnabled ||
      prevInitProps.notificationsEnabled !== nextInitProps.notificationsEnabled ||
      prevInitProps.siteIds.length !== nextInitProps.siteIds.length ||
      prevInitProps.siteIds.some((id, index) => id !== nextInitProps.siteIds[index]);

    if (hasInitPropsChanged) {
      previousInitPropsRef.current = {
        ...nextInitProps,
        siteIds: [...nextInitProps.siteIds]
      };

      initializeAlerts({
        graphClient: props.graphClient,
        context: props.context,
        siteIds: uniqueSortedSiteIds,
        alertTypesJson: props.alertTypesJson,
        userTargetingEnabled: !!props.userTargetingEnabled,
        notificationsEnabled: !!props.notificationsEnabled
      });
    }

    setIsInEditMode(EditModeDetector.isPageInEditMode(props.context));
  }, [
    props.graphClient,
    props.context,
    props.alertTypesJson,
    props.userTargetingEnabled,
    props.notificationsEnabled,
    (props.siteIds || []).join('|'),
    initializeAlerts
  ]);

  // Monitor for edit mode changes using MutationObserver with debouncing
  React.useEffect(() => {
    let debounceTimer: number | null = null;

    const checkEditMode = () => {
      const newEditMode = EditModeDetector.isPageInEditMode(props.context);
      setIsInEditMode(prevMode => {
        if (prevMode !== newEditMode) {
          logger.debug('Alerts', `Edit mode changed: ${prevMode} -> ${newEditMode}`);
        }
        return newEditMode;
      });
    };

    const debouncedCheckEditMode = () => {
      if (debounceTimer) {
        window.clearTimeout(debounceTimer);
      }
      debounceTimer = window.setTimeout(checkEditMode, 200);
    };

    const observer = new MutationObserver(debouncedCheckEditMode);

    observer.observe(document.body, {
      attributes: true,
      attributeFilter: ['class'],
      subtree: false
    });

    const observeCommandBar = () => {
      const commandBar = document.querySelector('[data-automation-id="pageCommandBarRegion"]');
      if (commandBar) {
        observer.observe(commandBar, {
          childList: true,
          subtree: true
        });
      }
    };

    observeCommandBar();

    const retryTimer = window.setTimeout(observeCommandBar, 1000);
    checkEditMode();
    return () => {
      if (debounceTimer) {
        window.clearTimeout(debounceTimer);
      }
      window.clearTimeout(retryTimer);
      observer.disconnect();
    };
  }, []);

  // Effect to reset index when alerts change
  React.useEffect(() => {
    if (alerts.length > 0 && currentIndex >= alerts.length) {
      setCurrentIndex(alerts.length - 1);
    } else if (alerts.length === 0) {
      setCurrentIndex(0);
    }
  }, [alerts, currentIndex]);

  // Carousel timer effect
  React.useEffect(() => {
    if (carouselEnabled && alerts.length > 1) {
      carouselTimer.current = window.setInterval(() => {
        setCurrentIndex(prevIndex => (prevIndex + 1) % alerts.length);
      }, carouselInterval);
    } else if (carouselTimer.current) {
      window.clearInterval(carouselTimer.current);
      carouselTimer.current = null;
    }

    // Cleanup
    return () => {
      if (carouselTimer.current) {
        window.clearInterval(carouselTimer.current);
      }
    };
  }, [carouselEnabled, carouselInterval, alerts.length]);

  React.useEffect(() => {
    const savedCarouselEnabled = storageService.current.getFromLocalStorage<boolean>('carouselEnabled');
    const savedCarouselInterval = storageService.current.getFromLocalStorage<number>('carouselInterval');

    if (savedCarouselEnabled !== null) {
      setCarouselEnabled(savedCarouselEnabled);
    }
    if (savedCarouselInterval && savedCarouselInterval >= CAROUSEL_CONFIG.MIN_INTERVAL && savedCarouselInterval <= CAROUSEL_CONFIG.MAX_INTERVAL) {
      setCarouselInterval(savedCarouselInterval);
    }

    // Listen for carousel settings changes from the settings panel
    const handleCarouselSettingsChange = (event: CustomEvent): void => {
      const { carouselEnabled: newEnabled, carouselInterval: newInterval } = event.detail;

      if (newEnabled !== undefined && newEnabled !== null) {
        setCarouselEnabled(newEnabled);
      }
      if (newInterval && newInterval >= CAROUSEL_CONFIG.MIN_INTERVAL && newInterval <= CAROUSEL_CONFIG.MAX_INTERVAL) {
        setCarouselInterval(newInterval);
      }
    };

    window.addEventListener('carouselSettingsChanged', handleCarouselSettingsChange as EventListener);

    return () => {
      window.removeEventListener('carouselSettingsChanged', handleCarouselSettingsChange as EventListener);
    };
  }, []);

  const handleSettingsChange = React.useCallback((settings: ISettingsData) => {
    if (props.onSettingsChange) {
      props.onSettingsChange(settings);
    }
    // The context will handle reloading alert types if they changed via its own logic
  }, [props.onSettingsChange]);

  // Carousel navigation with useCallback optimization
  const goToNext = React.useCallback(() => {
    setCurrentIndex((prevIndex) => (prevIndex + 1) % alerts.length);
  }, [alerts.length]);

  const goToPrevious = React.useCallback(() => {
    setCurrentIndex((prevIndex) => (prevIndex - 1 + alerts.length) % alerts.length);
  }, [alerts.length]);

  // Carousel pause functionality with useCallback optimization
  const handleMouseEnter = React.useCallback(() => {
    if (carouselTimer.current) {
      window.clearInterval(carouselTimer.current);
      carouselTimer.current = null;
    }
  }, []);

  const handleMouseLeave = React.useCallback(() => {
    if (carouselEnabled && alerts.length > 1) {
      carouselTimer.current = window.setInterval(() => {
        setCurrentIndex(prevIndex => (prevIndex + 1) % alerts.length);
      }, carouselInterval);
    }
  }, [carouselEnabled, alerts.length, carouselInterval]);

  if (isLoading) {
    return null; // Hide loading, let alerts load silently in the background
  }

  if (hasError) {
    return (
      <div 
        className={styles.errorContainer}
        style={{
          padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
          backgroundColor: tokens.colorNeutralBackground1,
          borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
          fontFamily: tokens.fontFamilyBase,
        }}
      >
        <MessageBar 
          intent="error"
          className={styles.errorMessageBar}
          style={{
            borderRadius: tokens.borderRadiusMedium,
            boxShadow: tokens.shadow4,
          }}
        >
          <MessageBarBody>
            <MessageBarTitle style={{ 
              color: tokens.colorPaletteRedForeground1,
              fontWeight: tokens.fontWeightSemibold,
              fontSize: tokens.fontSizeBase300
            }}>
              {strings.AlertsLoadErrorTitle}
            </MessageBarTitle>
            <div style={{
              marginTop: tokens.spacingVerticalXS,
              fontSize: tokens.fontSizeBase200,
              lineHeight: tokens.lineHeightBase200,
            }}>
              {errorMessage || strings.AlertsLoadErrorFallback}
            </div>
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
        <div
          className={styles.carousel}
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
        >
          <ErrorBoundary componentName="AlertItem" onError={(error) => logger.error('Alerts', 'Alert rendering failed', error)}>
            <AlertItem
              key={alerts[currentIndex].id}
              item={alerts[currentIndex]}
              remove={removeAlert}
              hideForever={hideAlertForever}
              alertType={alertTypes[alerts[currentIndex].AlertType] || defaultAlertType}
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
          graphClient={props.graphClient}
          context={props.context}
          onSettingsChange={handleSettingsChange}
        />
      )}
    </div>
  );
};

export default Alerts;
