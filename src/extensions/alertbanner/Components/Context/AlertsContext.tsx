import * as React from "react";
import { createContext, useReducer, useContext, useCallback } from "react";
import { IAlertType, AlertPriority, IPersonField, ITargetingRule, ContentType, TargetLanguage, NotificationType } from "../Alerts/IAlerts";
import { IAlertItem } from "../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import UserTargetingService from "../Services/UserTargetingService";
import { NotificationService } from "../Services/NotificationService";
import StorageService from "../Services/StorageService";
import { LanguageAwarenessService } from "../Services/LanguageAwarenessService";
import { logger } from "../Services/LoggerService";
import { validationService } from "../Services/ValidationService";
import { AlertTransformers } from "../Utils/AlertTransformers";
import { AlertFilters } from "../Utils/AlertFilters";

interface AlertsState {
  alerts: IAlertItem[];
  alertTypes: { [key: string]: IAlertType };
  isLoading: boolean;
  hasError: boolean;
  errorMessage?: string;
  userDismissedAlerts: string[];
  userHiddenAlerts: string[];
}

type AlertsAction =
  | { type: 'SET_ALERTS'; payload: IAlertItem[] }
  | { type: 'SET_ALERT_TYPES'; payload: { [key: string]: IAlertType } }
  | { type: 'SET_LOADING'; payload: boolean }
  | { type: 'SET_ERROR'; payload: { hasError: boolean; message?: string } }
  | { type: 'DISMISS_ALERT'; payload: string }
  | { type: 'HIDE_ALERT_FOREVER'; payload: string }
  | { type: 'SET_DISMISSED_ALERTS'; payload: string[] }
  | { type: 'SET_HIDDEN_ALERTS'; payload: string[] }
  | { type: 'BATCH_UPDATE'; payload: Partial<AlertsState> };

const initialState: AlertsState = {
  alerts: [],
  alertTypes: {},
  isLoading: true,
  hasError: false,
  errorMessage: undefined,
  userDismissedAlerts: [],
  userHiddenAlerts: []
};

const alertsReducer = (state: AlertsState, action: AlertsAction): AlertsState => {
  switch (action.type) {
    case 'SET_ALERTS':
      return { ...state, alerts: action.payload };

    case 'SET_ALERT_TYPES':
      return { ...state, alertTypes: action.payload };

    case 'SET_LOADING':
      return { ...state, isLoading: action.payload };

    case 'SET_ERROR':
      return {
        ...state,
        hasError: action.payload.hasError,
        errorMessage: action.payload.message
      };

    case 'DISMISS_ALERT':
      return {
        ...state,
        alerts: state.alerts.filter(alert => alert.id !== action.payload),
        userDismissedAlerts: [...state.userDismissedAlerts, action.payload]
      };

    case 'HIDE_ALERT_FOREVER':
      return {
        ...state,
        alerts: state.alerts.filter(alert => alert.id !== action.payload),
        userHiddenAlerts: [...state.userHiddenAlerts, action.payload]
      };

    case 'SET_DISMISSED_ALERTS':
      return { ...state, userDismissedAlerts: action.payload };

    case 'SET_HIDDEN_ALERTS':
      return { ...state, userHiddenAlerts: action.payload };

    case 'BATCH_UPDATE':
      return { ...state, ...action.payload };

    default:
      return state;
  }
};

// Create the context
interface AlertsContextProps {
  state: AlertsState;
  dispatch: React.Dispatch<AlertsAction>;
  removeAlert: (id: string) => void;
  hideAlertForever: (id: string) => void;
  initializeAlerts: (options: AlertsContextOptions) => Promise<void>;
  refreshAlerts: () => Promise<void>;
}

const AlertsContext = createContext<AlertsContextProps | undefined>(undefined);

export interface AlertsContextOptions {
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  siteIds: string[];
  alertTypesJson: string;
  userTargetingEnabled?: boolean;
  notificationsEnabled?: boolean;
}

export const AlertsProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [state, dispatch] = useReducer(alertsReducer, initialState);
  
  const storageService = React.useMemo(() => StorageService.getInstance(), []);

  const servicesRef = React.useRef<{
    graphClient?: MSGraphClientV3;
    userTargetingService?: UserTargetingService;
    notificationService?: NotificationService;
    languageAwarenessService?: LanguageAwarenessService;
    options?: AlertsContextOptions;
  }>({});

  // Load alert types from JSON
  const loadAlertTypes = useCallback((alertTypesJson: string) => {
    const endPerformanceTracking = logger.startPerformanceTracking('loadAlertTypes');

    try {
      // Validate JSON with security checks before parsing
      const validation = validationService.validateJson(alertTypesJson, 10);

      if (!validation.isValid) {
        logger.error("AlertsContext", "Invalid alert types JSON - security validation failed", {
          errors: validation.errors,
          jsonPreview: alertTypesJson.substring(0, 100)
        });
        dispatch({ type: 'SET_ALERT_TYPES', payload: {} });
        return;
      }

      // Use sanitized/validated JSON data
      const alertTypesData: IAlertType[] = validation.sanitizedValue;
      const alertTypesMap: { [key: string]: IAlertType } = {};

      alertTypesData.forEach((type) => {
        alertTypesMap[type.name] = type;
      });

      dispatch({ type: 'SET_ALERT_TYPES', payload: alertTypesMap });
      logger.info('AlertsContext', `Loaded ${alertTypesData.length} alert types (validated)`);

    } catch (error) {
      logger.error("AlertsContext", "Error processing alert types JSON", error, { alertTypesJson });
      dispatch({ type: 'SET_ALERT_TYPES', payload: {} });
    } finally {
      endPerformanceTracking();
    }
  }, []);

  // Map SharePoint item to alert using consolidated transformer
  const mapSharePointItemToAlert = useCallback((item: any, siteId: string): IAlertItem => {
    return AlertTransformers.mapSharePointItemToAlert(item, siteId, false);
  }, []);

  const alertCacheRef = React.useRef<Map<string, { alerts: IAlertItem[]; timestamp: number }>>(new Map());
  const pendingRequestsRef = React.useRef<Map<string, Promise<IAlertItem[]>>>(new Map());
  const CACHE_DURATION = 5 * 60 * 1000;

  // Normalize site ID to extract the site GUID for consistent deduplication
  const normalizeSiteId = useCallback((siteId: string): string => {
    // If it's in Graph format (hostname,siteGuid,webGuid), extract just the site GUID
    if (siteId.includes(',')) {
      const parts = siteId.split(',');
      // Return the middle part (site GUID)
      return parts.length >= 2 ? parts[1] : siteId;
    }

    // If it contains a colon (hostname:path format), we can't easily extract GUID
    // Return as-is and let Graph API handle it
    if (siteId.includes(':')) {
      return siteId;
    }

    // Otherwise, it's already a GUID - return as-is
    return siteId;
  }, []);

  const fetchAlerts = useCallback(async (siteId: string): Promise<IAlertItem[]> => {
    // Check for pending request first (request deduplication)
    const pendingRequest = pendingRequestsRef.current.get(siteId);
    if (pendingRequest) {
      logger.debug('AlertsContext', `Reusing pending request for site ${siteId}`);
      return pendingRequest;
    }

    // Check cache
    const cached = alertCacheRef.current.get(siteId);
    const now = Date.now();
    if (cached && (now - cached.timestamp) < CACHE_DURATION) {
      logger.debug('AlertsContext', `Using cached alerts for site ${siteId}`, {
        alertCount: cached.alerts.length,
        cacheAge: now - cached.timestamp
      });
      return cached.alerts;
    }

    const dateTimeNow = new Date().toISOString();

    // Create new request and store it in pending map
    const requestPromise = (async () => {
      try {
      // First, try to get list with custom fields
      let response;
      try {
        const filterQuery = `(fields/ScheduledStart le '${dateTimeNow}' or fields/ScheduledStart eq null) and (fields/ScheduledEnd ge '${dateTimeNow}' or fields/ScheduledEnd eq null) and (fields/ItemType ne 'template') and (fields/ItemType ne 'draft')`;
        if (!servicesRef.current.graphClient) throw new Error('GraphClient not initialized');
        response = await servicesRef.current.graphClient
          .api(`/sites/${siteId}/lists/Alerts/items`)
          .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
          .expand("fields($select=Title,AlertType,Description,ScheduledStart,ScheduledEnd,Priority,IsPinned,NotificationType,LinkUrl,LinkDescription,TargetSites,Status,ItemType,TargetLanguage,LanguageGroup,AvailableForAll,Metadata,Attachments,AttachmentFiles)")
          .filter(filterQuery)
          .orderby("fields/ScheduledStart desc")
          .top(25)
          .get();
      } catch (customFieldError) {
        logger.warn('AlertsContext', 'Custom fields not found, falling back to basic fields', customFieldError);
        if (!servicesRef.current.graphClient) throw new Error('GraphClient not initialized');
        response = await servicesRef.current.graphClient
          .api(`/sites/${siteId}/lists/Alerts/items`)
          .expand("fields($select=Title,Description,Created,Author,Attachments,AttachmentFiles)")
          .orderby("fields/Created desc")
          .top(25)
          .get();
      }

      let alerts = response.value.map((item: any) => mapSharePointItemToAlert(item, siteId));

      // Filter out templates, drafts, and auto-saved items using AlertFilters utility
      alerts = AlertFilters.excludeNonPublicAlerts(alerts);

        // Cache the results
        alertCacheRef.current.set(siteId, { alerts, timestamp: now });
        logger.info('AlertsContext', `Fetched and cached ${alerts.length} alerts for site ${siteId}`);

        return alerts;
      } catch (error) {
        logger.error('AlertsContext', `Error fetching alerts from site ${siteId}`, error, { siteId });
        return [];
      } finally {
        // Remove from pending requests map when done
        pendingRequestsRef.current.delete(siteId);
      }
    })();

    // Store the promise in pending requests
    pendingRequestsRef.current.set(siteId, requestPromise);
    return requestPromise;
  }, [mapSharePointItemToAlert]);

  // Sort alerts by priority
  const sortAlertsByPriority = useCallback((alertsToSort: IAlertItem[]): IAlertItem[] => {
    const priorityOrder: { [key in AlertPriority]: number } = {
      [AlertPriority.Critical]: 0,
      [AlertPriority.High]: 1,
      [AlertPriority.Medium]: 2,
      [AlertPriority.Low]: 3
    };

    return [...alertsToSort].sort((a, b) => {
      // First sort by pinned status
      if (a.isPinned && !b.isPinned) return -1;
      if (!a.isPinned && b.isPinned) return 1;

      // Then sort by priority
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    });
  }, []);

  // Remove duplicates using AlertFilters utility
  const removeDuplicateAlerts = useCallback((alertsToFilter: IAlertItem[]): IAlertItem[] => {
    return AlertFilters.removeDuplicates(alertsToFilter);
  }, []);

  // Check if alerts have changed
  const areAlertsDifferent = useCallback((newAlerts: IAlertItem[], cachedAlerts: IAlertItem[] | null): boolean => {
    if (!cachedAlerts) return true;
    if (newAlerts.length !== cachedAlerts.length) return true;

    // Create maps for faster comparison
    const newAlertsMap = new Map(newAlerts.map(alert => [alert.id, alert]));
    const cachedAlertsMap = new Map(cachedAlerts.map(alert => [alert.id, alert]));

    const arraysMatch = (a: string[] = [], b: string[] = []): boolean => {
      if (a.length !== b.length) {
        return false;
      }
      return a.every((value, index) => value === b[index]);
    };

    const attachmentsMatch = (
      a: NonNullable<IAlertItem['attachments']> = [],
      b: NonNullable<IAlertItem['attachments']> = []
    ): boolean => {
      if (a.length !== b.length) {
        return false;
      }

      return a.every((attachment, index) => {
        const other = b[index];
        return attachment.fileName === other.fileName &&
          attachment.serverRelativeUrl === other.serverRelativeUrl &&
          (attachment.size ?? null) === (other.size ?? null);
      });
    };

    // Check if all IDs match
    if (newAlerts.some(alert => !cachedAlertsMap.has(alert.id))) return true;

    // Check if any alert properties have changed
    for (const [id, newAlert] of newAlertsMap.entries()) {
      const cachedAlert = cachedAlertsMap.get(id);
      if (!cachedAlert) return true;

      if (
        newAlert.title !== cachedAlert.title ||
        newAlert.description !== cachedAlert.description ||
        newAlert.AlertType !== cachedAlert.AlertType ||
        newAlert.priority !== cachedAlert.priority ||
        newAlert.isPinned !== cachedAlert.isPinned ||
        newAlert.linkUrl !== cachedAlert.linkUrl ||
        newAlert.linkDescription !== cachedAlert.linkDescription ||
        newAlert.status !== cachedAlert.status ||
        newAlert.notificationType !== cachedAlert.notificationType ||
        newAlert.contentType !== cachedAlert.contentType ||
        newAlert.targetLanguage !== cachedAlert.targetLanguage ||
        newAlert.languageGroup !== cachedAlert.languageGroup ||
        newAlert.availableForAll !== cachedAlert.availableForAll ||
        newAlert.scheduledStart !== cachedAlert.scheduledStart ||
        newAlert.scheduledEnd !== cachedAlert.scheduledEnd ||
        JSON.stringify(newAlert.metadata ?? null) !== JSON.stringify(cachedAlert.metadata ?? null) ||
        !arraysMatch(newAlert.targetSites, cachedAlert.targetSites) ||
        !attachmentsMatch(newAlert.attachments, cachedAlert.attachments) ||
        JSON.stringify(newAlert.targetUsers ?? []) !== JSON.stringify(cachedAlert.targetUsers ?? [])
      ) {
        return true;
      }
    }

    return false;
  }, []);

  // Filter alerts based on user preferences using AlertFilters utility
  const filterAlerts = useCallback((alertsToFilter: IAlertItem[]): IAlertItem[] => {
    return AlertFilters.filterDismissedAndHidden(
      alertsToFilter,
      state.userDismissedAlerts,
      state.userHiddenAlerts
    );
  }, [state.userDismissedAlerts, state.userHiddenAlerts]);

  // Filter alerts by target sites using AlertFilters utility
  const filterAlertsByTargetSites = useCallback((alertsToFilter: IAlertItem[]): IAlertItem[] => {
    const scopedSiteIds = servicesRef.current.options?.siteIds || [];
    return AlertFilters.filterByTargetSites(alertsToFilter, scopedSiteIds);
  }, []);

  // Apply language-aware filtering to show appropriate language variants
  const applyLanguageAwareFiltering = useCallback(async (alertsToFilter: IAlertItem[]): Promise<IAlertItem[]> => {
    if (!servicesRef.current.languageAwarenessService) {
      return alertsToFilter; // No language service, return as-is
    }

    try {
      const userLanguage = await servicesRef.current.languageAwarenessService.getUserPreferredLanguage();
      const filteredAlerts = servicesRef.current.languageAwarenessService.filterAlertsForUser(alertsToFilter, userLanguage);
      
      logger.info('AlertsContext', `Applied language filtering: ${userLanguage}`, {
        originalCount: alertsToFilter.length,
        filteredCount: filteredAlerts.length,
        userLanguage
      });
      
      return filteredAlerts;
    } catch (error) {
      logger.warn('AlertsContext', 'Error applying language filtering, using all alerts', error);
      return alertsToFilter;
    }
  }, []);

  // Send notifications
  const sendNotifications = useCallback(async (alertsToNotify: IAlertItem[]): Promise<void> => {
    if (!servicesRef.current.options?.notificationsEnabled || alertsToNotify.length === 0 || !servicesRef.current.notificationService) return;

    for (const alert of alertsToNotify) {
      if (alert.notificationType !== NotificationType.Browser && alert.notificationType !== NotificationType.Both) {
        continue;
      }

      await servicesRef.current.notificationService.showInfo(`New alert: ${alert.title}`, 'Alert Notification');
    }
  }, []);

  // Initialize alerts
  const initializeAlerts = useCallback(async (initOptions: AlertsContextOptions): Promise<void> => {
    try {
      servicesRef.current.options = initOptions;
      servicesRef.current.graphClient = initOptions.graphClient;

      // Initialize services
      servicesRef.current.userTargetingService = UserTargetingService.getInstance(servicesRef.current.graphClient);
      servicesRef.current.notificationService = NotificationService.getInstance(initOptions.context);
      servicesRef.current.languageAwarenessService = new LanguageAwarenessService(servicesRef.current.graphClient, initOptions.context);

      alertCacheRef.current.clear();

      dispatch({ type: 'SET_LOADING', payload: true });

      // Initialize user targeting service first
      if (servicesRef.current.options.userTargetingEnabled) {
        await servicesRef.current.userTargetingService.initialize();
      }

      // Load alert types from JSON
      loadAlertTypes(servicesRef.current.options.alertTypesJson);

      // Get user's dismissed and hidden alerts - batch update to reduce re-renders
      if (servicesRef.current.options.userTargetingEnabled) {
        const dismissedAlerts = servicesRef.current.userTargetingService.getUserDismissedAlerts();
        const hiddenAlerts = servicesRef.current.userTargetingService.getUserHiddenAlerts();

        dispatch({ 
          type: 'BATCH_UPDATE', 
          payload: { 
            userDismissedAlerts: dismissedAlerts,
            userHiddenAlerts: hiddenAlerts
          } 
        });
      }

      // Fetch and process alerts
      await refreshAlerts();

    } catch (error) {
      logger.error("AlertsContext", "Error initializing alerts", error, {
        options: servicesRef.current.options
      });
      dispatch({
        type: 'SET_ERROR',
        payload: {
          hasError: true,
          message: "Failed to load alerts. Please try refreshing the page."
        }
      });
      dispatch({ type: 'SET_LOADING', payload: false });
    }
  }, [loadAlertTypes]);

  // Refresh alerts
  const refreshAlerts = useCallback(async (): Promise<void> => {
    if (!servicesRef.current.options || !servicesRef.current.graphClient) return;

    try {
      const allAlerts: IAlertItem[] = [];

      // Process only 3 sites at a time to avoid performance issues
      const batchSize = 3;
      const siteIds = servicesRef.current.options.siteIds || [];

      // Normalize site IDs and remove duplicates
      const normalizedSiteIds = siteIds.map(normalizeSiteId);
      const uniqueSiteIds = [...new Set(normalizedSiteIds)];

      logger.info('AlertsContext', 'Processing sites for alert refresh', {
        totalSiteIds: siteIds.length,
        uniqueSiteIds: uniqueSiteIds.length
      });

      for (let i = 0; i < uniqueSiteIds.length; i += batchSize) {
        const batch = uniqueSiteIds.slice(i, i + batchSize);
        const batchPromises = batch.map(siteId => fetchAlerts(siteId));
        const batchResults = await Promise.allSettled(batchPromises);

        batchResults.forEach((result, index) => {
          if (result.status === 'fulfilled') {
            logger.debug('AlertsContext', `Site returned alerts successfully`, {
              siteId: batch[index],
              alertCount: result.value.length
            });
            allAlerts.push(...result.value);
          } else {
            logger.warn('AlertsContext', `Site failed to return alerts`, {
              siteId: batch[index],
              error: result.reason
            });
          }
        });
      }

      // If no alerts were fetched, handle gracefully
      if (allAlerts.length === 0) {
        dispatch({ type: 'SET_ALERTS', payload: [] });
        dispatch({ type: 'SET_LOADING', payload: false });
        return;
      }

      // Remove duplicates
      const uniqueAlerts = removeDuplicateAlerts(allAlerts);

      // Compare with cached alerts
      const cachedAlerts = storageService.getFromLocalStorage<IAlertItem[]>('AllAlerts');
      const alertsAreDifferent = areAlertsDifferent(uniqueAlerts, cachedAlerts);

      // Update cache if needed
      if (alertsAreDifferent) {
        storageService.saveToLocalStorage('AllAlerts', uniqueAlerts);
      }

      // Get alerts to display
      let alertsToShow = alertsAreDifferent ? uniqueAlerts : cachedAlerts || [];

      alertsToShow = filterAlertsByTargetSites(alertsToShow);

      // Apply user targeting if enabled
      if (servicesRef.current.options.userTargetingEnabled && servicesRef.current.userTargetingService) {
        alertsToShow = await servicesRef.current.userTargetingService.filterAlertsForCurrentUser(alertsToShow);
      }

      // Apply language-aware filtering to show appropriate language variants
      alertsToShow = await applyLanguageAwareFiltering(alertsToShow);

      // Filter out hidden/dismissed alerts
      alertsToShow = filterAlerts(alertsToShow);

      // Limit the number of alerts to prevent performance issues
      if (alertsToShow.length > 20) {
        logger.warn('AlertsContext', `Limiting alerts to 20 for performance (found ${alertsToShow.length})`);
        alertsToShow = alertsToShow.slice(0, 20);
      }

      // Sort alerts by priority
      alertsToShow = sortAlertsByPriority(alertsToShow);

      // Send notifications for critical/high priority alerts if they're new
      if (servicesRef.current.options.notificationsEnabled && alertsAreDifferent) {
        const highPriorityAlerts = alertsToShow.filter(alert =>
          alert.priority === AlertPriority.Critical ||
          alert.priority === AlertPriority.High
        );

        // Only send notifications for the first 5 high priority alerts to avoid spamming
        if (highPriorityAlerts.length > 0) {
          sendNotifications(highPriorityAlerts.slice(0, 5));
        }
      }

      // Update state
      dispatch({ type: 'SET_ALERTS', payload: alertsToShow });
      dispatch({ type: 'SET_LOADING', payload: false });

    } catch (error) {
      logger.error('AlertsContext', 'Error refreshing alerts', error);
      dispatch({
        type: 'SET_ERROR',
        payload: {
          hasError: true,
          message: "Failed to refresh alerts. Please try again."
        }
      });
      dispatch({ type: 'SET_LOADING', payload: false });
    }
  }, [
    filterAlerts,
    filterAlertsByTargetSites,
    removeDuplicateAlerts,
    sortAlertsByPriority,
    areAlertsDifferent,
    sendNotifications,
    fetchAlerts,
    normalizeSiteId,
    applyLanguageAwareFiltering
  ]);

  // Handle removing an alert
  const removeAlert = useCallback((id: string): void => {
    dispatch({ type: 'DISMISS_ALERT', payload: id });

    // Add to user's dismissed alerts if targeting is enabled
    if (servicesRef.current.options?.userTargetingEnabled && servicesRef.current.userTargetingService) {
      servicesRef.current.userTargetingService.addUserDismissedAlert(id);
    }
  }, []);

  // Handle hiding an alert forever
  const hideAlertForever = useCallback((id: string): void => {
    dispatch({ type: 'HIDE_ALERT_FOREVER', payload: id });

    // Add to user's hidden alerts if targeting is enabled
    if (servicesRef.current.options?.userTargetingEnabled && servicesRef.current.userTargetingService) {
      servicesRef.current.userTargetingService.addUserHiddenAlert(id);
    }
  }, []);

  // The value we'll provide to consumers
  const value = React.useMemo(() => ({
    state,
    dispatch,
    removeAlert,
    hideAlertForever,
    initializeAlerts,
    refreshAlerts
  }), [state, removeAlert, hideAlertForever, initializeAlerts, refreshAlerts]);

  // Periodic cleanup of expired cache entries
  const cleanupExpiredCache = useCallback(() => {
    const now = Date.now();
    let cleanedCount = 0;

    alertCacheRef.current.forEach((value, key) => {
      if (now - value.timestamp > CACHE_DURATION) {
        alertCacheRef.current.delete(key);
        cleanedCount++;
      }
    });

    if (cleanedCount > 0) {
      logger.debug('AlertsContext', `Cleaned up ${cleanedCount} expired cache entries`);
    }
  }, []);

  // Set up periodic cache cleanup
  React.useEffect(() => {
    const cleanupInterval = setInterval(cleanupExpiredCache, CACHE_DURATION);
    return () => clearInterval(cleanupInterval);
  }, [cleanupExpiredCache]);

  const cleanup = useCallback(() => {
    logger.info('AlertsContext', 'Cleaning up AlertsContext resources');
    alertCacheRef.current.clear();
    pendingRequestsRef.current.clear();
    servicesRef.current = {};

    // Dispatch final cleanup state
    dispatch({ type: 'SET_ALERTS', payload: [] });
    dispatch({ type: 'SET_LOADING', payload: false });
    dispatch({ type: 'SET_ERROR', payload: { hasError: false } });
  }, []);

  React.useEffect(() => {
    return cleanup;
  }, [cleanup]);

  return (
    <AlertsContext.Provider value={value}>
      {children}
    </AlertsContext.Provider>
  );
};

export const useAlerts = () => {
  const context = useContext(AlertsContext);
  if (context === undefined) {
    throw new Error('useAlerts must be used within an AlertsProvider');
  }
  return context;
};
