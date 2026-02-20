import * as React from "react";
import { createContext, useReducer, useContext, useCallback } from "react";
import {
  IAlertType,
  AlertPriority,
  ContentType,
  TargetLanguage,
  NotificationType,
  IAlertItem,
  IGraphListItem,
  IPriorityColorConfig,
} from "../Alerts/IAlerts";
import { SharePointAlertService } from "../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import UserTargetingService from "../Services/UserTargetingService";
import { NotificationService } from "../Services/NotificationService";
import StorageService from "../Services/StorageService";
import { LanguageAwarenessService } from "../Services/LanguageAwarenessService";
import {
  DEFAULT_LANGUAGE_POLICY,
  ILanguagePolicy,
} from "../Services/LanguagePolicyService";
import { logger } from "../Services/LoggerService";
import { validationService } from "../Services/ValidationService";
import { AlertTransformers } from "../Utils/AlertTransformers";
import { AlertFilters } from "../Utils/AlertFilters";
import { LIST_NAMES, CACHE_CONFIG, API_CONFIG } from "../Utils/AppConstants";
import { JsonUtils } from "../Utils/JsonUtils";
import { SiteIdUtils } from "../Utils/SiteIdUtils";

export type AlertSortMode = "manual" | "priority" | "date" | "alphabetical";

interface AlertsState {
  alerts: IAlertItem[];
  alertTypes: { [key: string]: IAlertType };
  isLoading: boolean;
  hasError: boolean;
  errorMessage?: string;
  userDismissedAlerts: string[];
  userHiddenAlerts: string[];
  carouselEnabled: boolean;
  carouselInterval: number;
  priorityBorderColors: Record<AlertPriority, IPriorityColorConfig>;
  sortMode: AlertSortMode;
}

type AlertsAction =
  | { type: "SET_ALERTS"; payload: IAlertItem[] }
  | { type: "SET_ALERT_TYPES"; payload: { [key: string]: IAlertType } }
  | { type: "SET_LOADING"; payload: boolean }
  | { type: "SET_ERROR"; payload: { hasError: boolean; message?: string } }
  | { type: "DISMISS_ALERT"; payload: string }
  | { type: "HIDE_ALERT_FOREVER"; payload: string }
  | { type: "SET_DISMISSED_ALERTS"; payload: string[] }
  | { type: "SET_HIDDEN_ALERTS"; payload: string[] }
  | {
      type: "SET_CAROUSEL_SETTINGS";
      payload: { carouselEnabled?: boolean; carouselInterval?: number };
    }
  | {
      type: "SET_PRIORITY_BORDER_COLORS";
      payload: Record<AlertPriority, IPriorityColorConfig>;
    }
  | { type: "SET_SORT_MODE"; payload: AlertSortMode }
  | { type: "BATCH_UPDATE"; payload: Partial<AlertsState> };

const initialState: AlertsState = {
  alerts: [],
  alertTypes: {},
  isLoading: true,
  hasError: false,
  errorMessage: undefined,
  userDismissedAlerts: [],
  userHiddenAlerts: [],
  carouselEnabled: false,
  carouselInterval: 5000,
  priorityBorderColors: {
    [AlertPriority.Critical]: { borderColor: "#d13438" },
    [AlertPriority.High]: { borderColor: "#f7630c" },
    [AlertPriority.Medium]: { borderColor: "#0078d4" },
    [AlertPriority.Low]: { borderColor: "#107c10" },
  },
  sortMode: "priority",
};

const alertsReducer = (
  state: AlertsState,
  action: AlertsAction,
): AlertsState => {
  switch (action.type) {
    case "SET_ALERTS":
      return { ...state, alerts: action.payload };

    case "SET_ALERT_TYPES":
      return { ...state, alertTypes: action.payload };

    case "SET_LOADING":
      return { ...state, isLoading: action.payload };

    case "SET_ERROR":
      return {
        ...state,
        hasError: action.payload.hasError,
        errorMessage: action.payload.message,
      };

    case "DISMISS_ALERT":
      return {
        ...state,
        alerts: state.alerts.filter((alert) => alert.id !== action.payload),
        userDismissedAlerts: [...state.userDismissedAlerts, action.payload],
      };

    case "HIDE_ALERT_FOREVER":
      return {
        ...state,
        alerts: state.alerts.filter((alert) => alert.id !== action.payload),
        userHiddenAlerts: [...state.userHiddenAlerts, action.payload],
      };

    case "SET_DISMISSED_ALERTS":
      return { ...state, userDismissedAlerts: action.payload };

    case "SET_HIDDEN_ALERTS":
      return { ...state, userHiddenAlerts: action.payload };

    case "SET_CAROUSEL_SETTINGS":
      return {
        ...state,
        ...(action.payload.carouselEnabled !== undefined && {
          carouselEnabled: action.payload.carouselEnabled,
        }),
        ...(action.payload.carouselInterval !== undefined && {
          carouselInterval: action.payload.carouselInterval,
        }),
      };

    case "SET_PRIORITY_BORDER_COLORS":
      return {
        ...state,
        priorityBorderColors: action.payload,
      };

    case "SET_SORT_MODE":
      return {
        ...state,
        sortMode: action.payload,
      };

    case "BATCH_UPDATE":
      return { ...state, ...action.payload };

    default:
      return state;
  }
};

export interface AlertsDispatchContextProps {
  dispatch: React.Dispatch<AlertsAction>;
  removeAlert: (id: string) => void;
  hideAlertForever: (id: string) => void;
  initializeAlerts: (options: AlertsContextOptions) => Promise<void>;
  refreshAlerts: () => Promise<void>;
  updateCarouselSettings: (settings: {
    carouselEnabled?: boolean;
    carouselInterval?: number;
  }) => void;
  updatePriorityBorderColors: (
    colors: Record<AlertPriority, IPriorityColorConfig>,
  ) => void;
  updateSortMode: (mode: AlertSortMode) => void;
}

export const AlertsStateContext = createContext<AlertsState | undefined>(
  undefined,
);
export const AlertsDispatchContext = createContext<
  AlertsDispatchContextProps | undefined
>(undefined);

export interface AlertsContextOptions {
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  siteIds: string[];
  alertTypesJson: string;
  userTargetingEnabled?: boolean;
  notificationsEnabled?: boolean;
  enableTargetSite?: boolean;
}

export const AlertsProvider: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => {
  const [state, dispatch] = useReducer(alertsReducer, initialState);

  const stateRef = React.useRef(state);
  React.useEffect(() => {
    stateRef.current = state;
  }, [state]);

  const storageService = React.useMemo(() => StorageService.getInstance(), []);

  const servicesRef = React.useRef<{
    graphClient?: MSGraphClientV3;
    userTargetingService?: UserTargetingService;
    notificationService?: NotificationService;
    languageAwarenessService?: LanguageAwarenessService;
    sharePointAlertService?: SharePointAlertService;
    languagePolicy?: ILanguagePolicy;
    options?: AlertsContextOptions;
  }>({});

  const loadAlertTypes = useCallback((alertTypesJson: string) => {
    const endPerformanceTracking =
      logger.startPerformanceTracking("loadAlertTypes");

    try {
      const validation = validationService.validateJson(alertTypesJson, 10);

      if (!validation.isValid) {
        logger.error(
          "AlertsContext",
          "Invalid alert types JSON - security validation failed",
          {
            errors: validation.errors,
            jsonPreview: alertTypesJson.substring(0, 100),
          },
        );
        dispatch({ type: "SET_ALERT_TYPES", payload: {} });
        return;
      }

      const alertTypesData: IAlertType[] =
        validation.sanitizedValue as IAlertType[];
      const alertTypesMap: { [key: string]: IAlertType } = {};

      alertTypesData.forEach((type) => {
        alertTypesMap[type.name] = type;
      });

      dispatch({ type: "SET_ALERT_TYPES", payload: alertTypesMap });
      logger.info(
        "AlertsContext",
        `Loaded ${alertTypesData.length} alert types (validated)`,
      );
    } catch (error) {
      logger.error(
        "AlertsContext",
        "Error processing alert types JSON",
        error,
        { alertTypesJson },
      );
      dispatch({ type: "SET_ALERT_TYPES", payload: {} });
    } finally {
      endPerformanceTracking();
    }
  }, []);

  const loadAlertTypesFromList = useCallback(async (): Promise<boolean> => {
    const endPerformanceTracking = logger.startPerformanceTracking(
      "loadAlertTypesFromList",
    );

    try {
      if (!servicesRef.current.sharePointAlertService) {
        return false;
      }

      const alertTypesData =
        await servicesRef.current.sharePointAlertService.getAlertTypes();
      if (!alertTypesData || alertTypesData.length === 0) {
        return false;
      }

      const alertTypesMap: { [key: string]: IAlertType } = {};
      alertTypesData.forEach((type) => {
        if (type?.name) {
          alertTypesMap[type.name] = type;
        }
      });

      dispatch({ type: "SET_ALERT_TYPES", payload: alertTypesMap });
      logger.info(
        "AlertsContext",
        `Loaded ${alertTypesData.length} alert types from SharePoint list`,
      );
      return true;
    } catch (error) {
      logger.warn(
        "AlertsContext",
        "Failed to load alert types from SharePoint list",
        error,
      );
      return false;
    } finally {
      endPerformanceTracking();
    }
  }, []);

  const mapSharePointItemToAlert = useCallback(
    (item: IGraphListItem, siteId: string): IAlertItem => {
      return AlertTransformers.mapSharePointItemToAlert(item, siteId, false);
    },
    [],
  );

  const alertCacheRef = React.useRef<
    Map<string, { alerts: IAlertItem[]; timestamp: number }>
  >(new Map());
  const pendingRequestsRef = React.useRef<Map<string, Promise<IAlertItem[]>>>(
    new Map(),
  );
  const CACHE_DURATION = CACHE_CONFIG.ALERTS_CACHE_DURATION;

  const createSiteDedupKey = useCallback((siteId: string): string => {
    return SiteIdUtils.createDedupKey(siteId);
  }, []);

  const buildAlertsCacheKey = useCallback(
    (siteIds: string[]): string => {
      const userLogin =
        servicesRef.current.options?.context?.pageContext?.user?.loginName ||
        "anonymous";
      const siteScope =
        siteIds
          .map((siteId) => createSiteDedupKey(siteId))
          .filter((siteId) => siteId.length > 0)
          .sort()
          .join("|") || "none";

      return `AllAlerts:${userLogin}:${siteScope}`;
    },
    [createSiteDedupKey],
  );

  const createAlertSignature = useCallback((alert: IAlertItem): string => {
    const signaturePayload = {
      id: alert.id,
      title: alert.title,
      description: alert.description,
      type: alert.AlertType,
      priority: alert.priority,
      pinned: alert.isPinned,
      linkUrl: alert.linkUrl,
      linkDescription: alert.linkDescription,
      status: alert.status,
      notificationType: alert.notificationType,
      contentType: alert.contentType,
      targetLanguage: alert.targetLanguage,
      languageGroup: alert.languageGroup,
      availableForAll: alert.availableForAll,
      scheduledStart: alert.scheduledStart,
      scheduledEnd: alert.scheduledEnd,
      metadata: alert.metadata ?? null,
      targetSites: alert.targetSites ?? [],
      attachments: (alert.attachments || []).map((attachment) => ({
        fileName: attachment.fileName,
        url: attachment.serverRelativeUrl,
        size: attachment.size ?? null,
      })),
      targetUsers: alert.targetUsers ?? [],
    };

    return JsonUtils.safeStringify(signaturePayload) || "";
  }, []);

  const fetchAlerts = useCallback(
    async (siteId: string): Promise<IAlertItem[]> => {
      const pendingRequest = pendingRequestsRef.current.get(siteId);
      if (pendingRequest) {
        logger.debug(
          "AlertsContext",
          `Reusing pending request for site ${siteId}`,
        );
        return pendingRequest;
      }

      const cached = alertCacheRef.current.get(siteId);
      const now = Date.now();
      if (cached && now - cached.timestamp < CACHE_DURATION) {
        logger.debug(
          "AlertsContext",
          `Using cached alerts for site ${siteId}`,
          {
            alertCount: cached.alerts.length,
            cacheAge: now - cached.timestamp,
          },
        );
        return cached.alerts;
      }

      const dateTimeNow = new Date().toISOString();

      const requestPromise = (async () => {
        try {
          if (!servicesRef.current.sharePointAlertService) {
            throw new Error("SharePointAlertService not initialized");
          }

          let alerts =
            await servicesRef.current.sharePointAlertService.getActiveAlerts(
              siteId,
            );

          // Filter out templates, drafts, and auto-saved items is already handled by SharePointAlertService.getActiveAlerts

          alertCacheRef.current.set(siteId, { alerts, timestamp: now });
          logger.info(
            "AlertsContext",
            `Fetched and cached ${alerts.length} alerts for site ${siteId}`,
          );

          return alerts;
        } catch (error) {
          logger.error(
            "AlertsContext",
            `Error fetching alerts from site ${siteId}`,
            error,
            { siteId },
          );
          return [];
        } finally {
          pendingRequestsRef.current.delete(siteId);
        }
      })();

      pendingRequestsRef.current.set(siteId, requestPromise);
      return requestPromise;
    },
    [mapSharePointItemToAlert],
  );

  const sortAlerts = useCallback(
    (
      alertsToSort: IAlertItem[],
      mode: AlertSortMode = stateRef.current.sortMode,
    ): IAlertItem[] => {
      const priorityOrder: { [key in AlertPriority]: number } = {
        [AlertPriority.Critical]: 0,
        [AlertPriority.High]: 1,
        [AlertPriority.Medium]: 2,
        [AlertPriority.Low]: 3,
      };

      return [...alertsToSort].sort((a, b) => {
        if (a.isPinned && !b.isPinned) return -1;
        if (!a.isPinned && b.isPinned) return 1;

        switch (mode) {
          case "manual":
            const orderA =
              typeof a.sortOrder === "number"
                ? a.sortOrder
                : Number.MAX_SAFE_INTEGER;
            const orderB =
              typeof b.sortOrder === "number"
                ? b.sortOrder
                : Number.MAX_SAFE_INTEGER;
            if (orderA !== orderB) {
              return orderA - orderB;
            }
            return priorityOrder[a.priority] - priorityOrder[b.priority];

          case "date":
            return (
              new Date(b.createdDate).getTime() -
              new Date(a.createdDate).getTime()
            );

          case "alphabetical":
            return a.title.localeCompare(b.title);

          case "priority":
          default:
            // Sort by priority (Critical > High > Medium > Low)
            return priorityOrder[a.priority] - priorityOrder[b.priority];
        }
      });
    },
    [],
  );

  const sortAlertsByPriority = sortAlerts;

  const removeDuplicateAlerts = useCallback(
    (alertsToFilter: IAlertItem[]): IAlertItem[] => {
      return AlertFilters.removeDuplicates(alertsToFilter);
    },
    [],
  );

  const areAlertsDifferent = useCallback(
    (newAlerts: IAlertItem[], cachedAlerts: IAlertItem[] | null): boolean => {
      if (!cachedAlerts) return true;
      if (newAlerts.length !== cachedAlerts.length) return true;

      const cachedMap = new Map(cachedAlerts.map((a) => [a.id, a]));

      for (const newAlert of newAlerts) {
        const cachedAlert = cachedMap.get(newAlert.id);
        if (!cachedAlert) {
          return true;
        }

        if (newAlert.modified && cachedAlert.modified) {
          if (newAlert.modified !== cachedAlert.modified) {
            return true;
          }
          continue;
        }

        const cachedSignature = createAlertSignature(cachedAlert);
        const currentSignature = createAlertSignature(newAlert);

        if (currentSignature !== cachedSignature) {
          return true;
        }
      }

      return false;
    },
    [createAlertSignature],
  );

  const filterAlerts = useCallback(
    (alertsToFilter: IAlertItem[]): IAlertItem[] => {
      return AlertFilters.filterDismissedAndHidden(
        alertsToFilter,
        state.userDismissedAlerts,
        state.userHiddenAlerts,
      );
    },
    [state.userDismissedAlerts, state.userHiddenAlerts],
  );

  const filterAlertsByTargetSites = useCallback(
    (alertsToFilter: IAlertItem[]): IAlertItem[] => {
      const scopedSiteIds = servicesRef.current.options?.siteIds || [];
      return AlertFilters.filterByTargetSites(alertsToFilter, scopedSiteIds);
    },
    [],
  );

  const applyLanguageAwareFiltering = useCallback(
    async (alertsToFilter: IAlertItem[]): Promise<IAlertItem[]> => {
      if (!servicesRef.current.languageAwarenessService) {
        return alertsToFilter; // No language service, return as-is
      }

      try {
        const detectionResult =
          await servicesRef.current.languageAwarenessService.getUserPreferredLanguageWithSource();
        const userPreferences =
          servicesRef.current.languageAwarenessService.getUserLanguagePreferences();

        const filteredAlerts =
          servicesRef.current.languageAwarenessService.filterAlertsForUser(
            alertsToFilter,
            detectionResult.language,
            servicesRef.current.languagePolicy || DEFAULT_LANGUAGE_POLICY,
          );

        logger.info(
          "AlertsContext",
          `Applied language filtering: ${detectionResult.language} (from ${detectionResult.source})`,
          {
            originalCount: alertsToFilter.length,
            filteredCount: filteredAlerts.length,
            userLanguage: detectionResult.language,
            languageSource: detectionResult.source,
            allPreferences: userPreferences,
          },
        );

        return filteredAlerts;
      } catch (error) {
        logger.warn(
          "AlertsContext",
          "Error applying language filtering, using all alerts",
          error,
        );
        return alertsToFilter;
      }
    },
    [],
  );

  const sendNotifications = useCallback(
    async (alertsToNotify: IAlertItem[]): Promise<void> => {
      if (
        !servicesRef.current.options?.notificationsEnabled ||
        alertsToNotify.length === 0 ||
        !servicesRef.current.notificationService
      )
        return;

      for (const alert of alertsToNotify) {
        if (
          alert.notificationType !== NotificationType.Browser &&
          alert.notificationType !== NotificationType.Both
        ) {
          continue;
        }

        await servicesRef.current.notificationService.showInfo(
          `New alert: ${alert.title}`,
          "Alert Notification",
        );
      }
    },
    [],
  );

  const initializeAlerts = useCallback(
    async (initOptions: AlertsContextOptions): Promise<void> => {
      const initStartTime = performance.now();

      try {
        servicesRef.current.options = initOptions;
        servicesRef.current.graphClient = initOptions.graphClient;

        servicesRef.current.userTargetingService =
          UserTargetingService.getInstance(
            servicesRef.current.graphClient,
            initOptions.context,
          );
        servicesRef.current.notificationService =
          NotificationService.getInstance(initOptions.context);
        servicesRef.current.languageAwarenessService =
          new LanguageAwarenessService(
            servicesRef.current.graphClient,
            initOptions.context,
          );
        servicesRef.current.sharePointAlertService = new SharePointAlertService(
          servicesRef.current.graphClient,
          initOptions.context,
        );

        alertCacheRef.current.clear();

        const siteIds = initOptions.siteIds || [];
        const dedupMap = new Map<string, string>();
        siteIds.forEach((siteId) => {
          const dedupKey = createSiteDedupKey(siteId);
          if (dedupKey && !dedupMap.has(dedupKey)) {
            dedupMap.set(dedupKey, siteId);
          }
        });
        const uniqueSiteIds = Array.from(dedupMap.values());
        const alertsCacheKey = buildAlertsCacheKey(uniqueSiteIds);

        const cachedAlerts =
          storageService.getFromLocalStorage<IAlertItem[]>(alertsCacheKey);
        const hasCachedAlerts = cachedAlerts && cachedAlerts.length > 0;

        // If we have cached alerts, start with them but keep loading state
        if (hasCachedAlerts) {
          let alertsToShow = filterAlertsByTargetSites(cachedAlerts);
          alertsToShow = filterAlerts(alertsToShow);
          alertsToShow = sortAlertsByPriority(alertsToShow);

          dispatch({ type: "SET_ALERTS", payload: alertsToShow });
          // Stay in loading state - we'll transition smoothly when fresh data arrives
          dispatch({ type: "SET_LOADING", payload: true });

          logger.info("AlertsContext", "Using cached alerts as placeholder", {
            cachedCount: cachedAlerts.length,
            renderTimeMs: Math.round(performance.now() - initStartTime),
          });
        } else {
          dispatch({ type: "SET_LOADING", payload: true });
        }

        const [languagePolicy, alertTypesLoaded, dismissedHiddenAlerts] =
          await Promise.all([
            servicesRef.current.sharePointAlertService
              .getLanguagePolicy()
              .catch(() => DEFAULT_LANGUAGE_POLICY),

            loadAlertTypesFromList(),

            (async () => {
              if (
                !servicesRef.current.options?.userTargetingEnabled ||
                !servicesRef.current.userTargetingService
              ) {
                return { dismissed: [], hidden: [] };
              }
              const initPromise =
                servicesRef.current.userTargetingService.initialize();
              await initPromise;
              return {
                dismissed:
                  servicesRef.current.userTargetingService.getUserDismissedAlerts(),
                hidden:
                  servicesRef.current.userTargetingService.getUserHiddenAlerts(),
              };
            })(),
          ]);

        servicesRef.current.languagePolicy = languagePolicy;

        if (!alertTypesLoaded && servicesRef.current.options) {
          loadAlertTypes(servicesRef.current.options.alertTypesJson);
        }

        dispatch({
          type: "BATCH_UPDATE",
          payload: {
            userDismissedAlerts: dismissedHiddenAlerts.dismissed,
            userHiddenAlerts: dismissedHiddenAlerts.hidden,
          },
        });

        await refreshAlerts();

        logger.info("AlertsContext", "Initialization complete", {
          totalTimeMs: Math.round(performance.now() - initStartTime),
          hadCachedData: hasCachedAlerts,
        });
      } catch (error) {
        logger.error("AlertsContext", "Error initializing alerts", error, {
          options: servicesRef.current.options,
        });
        dispatch({
          type: "SET_ERROR",
          payload: {
            hasError: true,
            message: "Failed to load alerts. Please try refreshing the page.",
          },
        });
        dispatch({ type: "SET_LOADING", payload: false });
      }
    },
    [
      loadAlertTypes,
      filterAlerts,
      filterAlertsByTargetSites,
      sortAlertsByPriority,
      buildAlertsCacheKey,
      createSiteDedupKey,
    ],
  );

  const refreshAlerts = useCallback(async (): Promise<void> => {
    if (!servicesRef.current.options || !servicesRef.current.graphClient)
      return;

    const refreshStartTime = performance.now();

    try {
      const allAlerts: IAlertItem[] = [];

      const batchSize = 5;
      const siteIds = servicesRef.current.options.siteIds || [];

      const dedupMap = new Map<string, string>();
      (siteIds || []).forEach((siteId) => {
        const dedupKey = createSiteDedupKey(siteId);
        if (dedupKey && !dedupMap.has(dedupKey)) {
          dedupMap.set(dedupKey, siteId);
        }
      });

      const uniqueSiteIds = Array.from(dedupMap.values());
      const alertsCacheKey = buildAlertsCacheKey(uniqueSiteIds);

      logger.info("AlertsContext", "Processing sites for alert refresh", {
        totalSiteIds: siteIds.length,
        uniqueSiteIds: uniqueSiteIds.length,
      });

      for (let i = 0; i < uniqueSiteIds.length; i += batchSize) {
        const batch = uniqueSiteIds.slice(i, i + batchSize);
        const batchPromises = batch.map((siteId) => fetchAlerts(siteId));
        const batchResults = await Promise.allSettled(batchPromises);

        batchResults.forEach((result, index) => {
          if (result.status === "fulfilled") {
            logger.debug("AlertsContext", `Site returned alerts successfully`, {
              siteId: batch[index],
              alertCount: result.value.length,
            });
            allAlerts.push(...result.value);
          } else {
            logger.warn("AlertsContext", `Site failed to return alerts`, {
              siteId: batch[index],
              error: result.reason,
            });
          }
        });
      }

      if (allAlerts.length === 0) {
        dispatch({ type: "SET_ALERTS", payload: [] });
        dispatch({ type: "SET_LOADING", payload: false });
        return;
      }

      const uniqueAlerts = removeDuplicateAlerts(allAlerts);

      const cachedAlerts =
        storageService.getFromLocalStorage<IAlertItem[]>(alertsCacheKey);
      const alertsAreDifferent = areAlertsDifferent(uniqueAlerts, cachedAlerts);

      if (alertsAreDifferent) {
        storageService.saveToLocalStorage(alertsCacheKey, uniqueAlerts);
      }

      let alertsToShow = alertsAreDifferent ? uniqueAlerts : cachedAlerts || [];

      alertsToShow = filterAlertsByTargetSites(alertsToShow);
      alertsToShow = filterAlerts(alertsToShow);

      if (
        servicesRef.current.options.userTargetingEnabled &&
        servicesRef.current.userTargetingService
      ) {
        try {
          const targetingPromise =
            servicesRef.current.userTargetingService.filterAlertsForCurrentUser(
              alertsToShow,
            );

          alertsToShow = await Promise.race([
            targetingPromise,
            new Promise<IAlertItem[]>((resolve) =>
              setTimeout(() => {
                logger.warn(
                  "AlertsContext",
                  "User targeting timed out, showing all alerts",
                );
                resolve(alertsToShow);
              }, 2000),
            ),
          ]);
        } catch (targetingError) {
          logger.warn(
            "AlertsContext",
            "User targeting failed, showing all alerts",
            targetingError,
          );
        }
      }

      alertsToShow = await applyLanguageAwareFiltering(alertsToShow);

      if (alertsToShow.length > 20) {
        logger.warn(
          "AlertsContext",
          `Limiting alerts to 20 for performance (found ${alertsToShow.length})`,
        );
        alertsToShow = alertsToShow.slice(0, 20);
      }

      alertsToShow = sortAlertsByPriority(alertsToShow);

      // Send notifications for critical/high priority alerts if they're new
      if (
        servicesRef.current.options.notificationsEnabled &&
        alertsAreDifferent
      ) {
        const highPriorityAlerts = alertsToShow.filter((alert) =>
          [AlertPriority.Critical, AlertPriority.High].includes(alert.priority),
        );

        // Only send notifications for the first 5 high priority alerts to avoid spamming
        if (highPriorityAlerts.length > 0) {
          sendNotifications(highPriorityAlerts.slice(0, 5));
        }
      }

      const currentAlerts = stateRef.current.alerts;
      const alertsActuallyChanged = areAlertsDifferent(
        alertsToShow,
        currentAlerts,
      );

      if (alertsActuallyChanged || stateRef.current.isLoading) {
        dispatch({
          type: "BATCH_UPDATE",
          payload: { alerts: alertsToShow, isLoading: false },
        });

        logger.info("AlertsContext", "Alert refresh complete - updated", {
          refreshTimeMs: Math.round(performance.now() - refreshStartTime),
          alertsCount: alertsToShow.length,
          changed: alertsActuallyChanged,
        });
      } else {
        dispatch({ type: "SET_LOADING", payload: false });

        logger.info("AlertsContext", "Alert refresh complete - no change", {
          refreshTimeMs: Math.round(performance.now() - refreshStartTime),
          alertsCount: alertsToShow.length,
        });
      }
    } catch (error) {
      logger.error("AlertsContext", "Error refreshing alerts", error);
      dispatch({
        type: "SET_ERROR",
        payload: {
          hasError: true,
          message: "Failed to refresh alerts. Please try again.",
        },
      });
      dispatch({ type: "SET_LOADING", payload: false });
    }
  }, [
    filterAlerts,
    filterAlertsByTargetSites,
    removeDuplicateAlerts,
    sortAlertsByPriority,
    areAlertsDifferent,
    sendNotifications,
    fetchAlerts,
    createSiteDedupKey,
    buildAlertsCacheKey,
    applyLanguageAwareFiltering,
  ]);

  const removeAlert = useCallback((id: string): void => {
    dispatch({ type: "DISMISS_ALERT", payload: id });

    if (
      servicesRef.current.options?.userTargetingEnabled &&
      servicesRef.current.userTargetingService
    ) {
      servicesRef.current.userTargetingService.addUserDismissedAlert(id);
    }
  }, []);

  const hideAlertForever = useCallback((id: string): void => {
    dispatch({ type: "HIDE_ALERT_FOREVER", payload: id });

    if (
      servicesRef.current.options?.userTargetingEnabled &&
      servicesRef.current.userTargetingService
    ) {
      servicesRef.current.userTargetingService.addUserHiddenAlert(id);
    }
  }, []);

  const updateCarouselSettings = useCallback(
    (settings: { carouselEnabled?: boolean; carouselInterval?: number }) => {
      dispatch({ type: "SET_CAROUSEL_SETTINGS", payload: settings });
    },
    [],
  );

  const updatePriorityBorderColors = useCallback(
    (colors: Record<AlertPriority, IPriorityColorConfig>) => {
      dispatch({ type: "SET_PRIORITY_BORDER_COLORS", payload: colors });
    },
    [],
  );

  const updateSortMode = useCallback(
    (mode: AlertSortMode) => {
      dispatch({ type: "SET_SORT_MODE", payload: mode });

      const currentAlerts = stateRef.current.alerts;
      if (currentAlerts.length > 0) {
        const reSortedAlerts = sortAlerts(currentAlerts, mode);
        dispatch({ type: "SET_ALERTS", payload: reSortedAlerts });
      }

      try {
        localStorage.setItem("AlertBanner_SortMode", mode);
      } catch {
      }
    },
    [sortAlerts],
  );

  React.useEffect(() => {
    try {
      const savedMode = localStorage.getItem(
        "AlertBanner_SortMode",
      ) as AlertSortMode | null;
      if (
        savedMode &&
        ["manual", "priority", "date", "alphabetical"].includes(savedMode)
      ) {
        dispatch({ type: "SET_SORT_MODE", payload: savedMode });
      }
    } catch {
    }
  }, []);

  const value = React.useMemo(
    () => ({
      state,
      dispatch,
      removeAlert,
      hideAlertForever,
      initializeAlerts,
      refreshAlerts,
      updateCarouselSettings,
      updatePriorityBorderColors,
      updateSortMode,
    }),
    [
      state,
      removeAlert,
      hideAlertForever,
      initializeAlerts,
      refreshAlerts,
      updateCarouselSettings,
      updatePriorityBorderColors,
      updateSortMode,
    ],
  );

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
      logger.debug(
        "AlertsContext",
        `Cleaned up ${cleanedCount} expired cache entries`,
      );
    }
  }, []);

  React.useEffect(() => {
    const cleanupInterval = setInterval(cleanupExpiredCache, CACHE_DURATION);
    return () => clearInterval(cleanupInterval);
  }, [cleanupExpiredCache]);

  const lastCrossTabRefreshRef = React.useRef<number>(0);
  const crossTabDebounceRef = React.useRef<number | null>(null);

  React.useEffect(() => {
    const storage = StorageService.getInstance();
    const CROSS_TAB_DEBOUNCE_MS = 2000; // Minimum time between cross-tab refreshes

    const cleanup = storage.initCrossTabSync(() => {
      const now = Date.now();
      const timeSinceLastRefresh = now - lastCrossTabRefreshRef.current;

      if (crossTabDebounceRef.current) {
        window.clearTimeout(crossTabDebounceRef.current);
        crossTabDebounceRef.current = null;
      }

      if (timeSinceLastRefresh < CROSS_TAB_DEBOUNCE_MS) {
        const delay = CROSS_TAB_DEBOUNCE_MS - timeSinceLastRefresh;
        logger.debug(
          "AlertsContext",
          `Cross-tab change detected, debouncing refresh for ${delay}ms`,
        );
        crossTabDebounceRef.current = window.setTimeout(() => {
          lastCrossTabRefreshRef.current = Date.now();
          refreshAlerts();
        }, delay);
        return;
      }

      logger.info(
        "AlertsContext",
        "Cross-tab storage change detected, refreshing state",
      );
      lastCrossTabRefreshRef.current = now;
      refreshAlerts();
    });

    return () => {
      if (crossTabDebounceRef.current) {
        window.clearTimeout(crossTabDebounceRef.current);
      }
      cleanup();
    };
  }, [refreshAlerts]);

  const cleanup = useCallback(() => {
    logger.info("AlertsContext", "Cleaning up AlertsContext resources");
    alertCacheRef.current.clear();
    pendingRequestsRef.current.clear();
    servicesRef.current = {};

    dispatch({
      type: "BATCH_UPDATE",
      payload: {
        alerts: [],
        isLoading: false,
        hasError: false,
      },
    });
  }, []);

  React.useEffect(() => {
    return cleanup;
  }, [cleanup]);

  const dispatchValue = React.useMemo(
    () => ({
      dispatch,
      removeAlert,
      hideAlertForever,
      initializeAlerts,
      refreshAlerts,
      updateCarouselSettings,
      updatePriorityBorderColors,
      updateSortMode,
    }),
    [
      dispatch,
      removeAlert,
      hideAlertForever,
      initializeAlerts,
      refreshAlerts,
      updateCarouselSettings,
      updatePriorityBorderColors,
      updateSortMode,
    ],
  );

  return (
    <AlertsDispatchContext.Provider value={dispatchValue}>
      <AlertsStateContext.Provider value={state}>
        {children}
      </AlertsStateContext.Provider>
    </AlertsDispatchContext.Provider>
  );
};

export const useAlertsState = () => {
  const context = useContext(AlertsStateContext);
  if (context === undefined) {
    throw new Error("useAlertsState must be used within an AlertsProvider");
  }
  return context;
};

export const useAlertsDispatch = () => {
  const context = useContext(AlertsDispatchContext);
  if (context === undefined) {
    throw new Error("useAlertsDispatch must be used within an AlertsProvider");
  }
  return context;
};
