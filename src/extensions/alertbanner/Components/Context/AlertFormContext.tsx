import * as React from "react";
import {
  AlertPriority,
  NotificationType,
  IAlertType,
  ContentType,
  TargetLanguage,
  IPersonField,
  IAlertItem,
  TranslationStatus,
} from "../Alerts/IAlerts";
import {
  ILanguageContent,
  ISupportedLanguage,
  LanguageAwarenessService,
} from "../Services/LanguageAwarenessService";
import { ILanguagePolicy, DEFAULT_LANGUAGE_POLICY } from "../Services/LanguagePolicyService";
import { SiteContextDetector, ISiteValidationResult } from "../Utils/SiteContextDetector";
import { SharePointAlertService } from "../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IFormErrors } from "../Settings/Tabs/SharedTypes";

// =============================================================================
// TYPES
// =============================================================================

/**
 * New alert form data structure
 */
export interface INewAlert {
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl: string;
  linkDescription: string;
  targetSites: string[];
  scheduledStart?: Date;
  scheduledEnd?: Date;
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageContent: ILanguageContent[];
  languageGroup?: string;
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
}

/**
 * Wizard step for create mode
 */
export type CreateWizardStep = "content" | "audience" | "publish";

/**
 * Complete state interface for the alert form
 */
export interface IAlertFormState {
  // Form data
  newAlert: INewAlert;
  errors: IFormErrors;

  // UI state
  isCreatingAlert: boolean;
  showPreview: boolean;
  showTemplates: boolean;
  createStep: CreateWizardStep;
  currentEntryMode: "scratch" | "templates" | "drafts" | "previous";

  // Progress and results
  creationProgress: ISiteValidationResult[];
  lastCreateAttemptFailed: boolean;

  // Auto-save state
  autoSaveDraftId: string | null;
  lastAutoSave: Date | null;
  autoSaveStatus: "idle" | "pending" | "saving" | "saved" | "error";
  isAutoSaving: boolean;

  // Multi-language support
  useMultiLanguage: boolean;
  supportedLanguages: ISupportedLanguage[];
  languagePolicy: ILanguagePolicy;

  // Drafts and previous alerts
  drafts: IAlertItem[];
  previousAlerts: IAlertItem[];

  // Copilot state
  copilotAvailability: "unknown" | "available" | "unavailable";

  // Dependencies (injected via provider)
  alertTypes: IAlertType[];
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  copilotEnabled: boolean;

}

// =============================================================================
// INITIAL STATE
// =============================================================================

export const createInitialNewAlert = (alertTypes: IAlertType[]): INewAlert => ({
  title: "",
  description: "",
  AlertType: alertTypes.length > 0 ? alertTypes[0].name : "",
  priority: AlertPriority.Medium,
  isPinned: false,
  notificationType: NotificationType.Browser,
  linkUrl: "",
  linkDescription: "",
  targetSites: [],
  scheduledStart: undefined,
  scheduledEnd: undefined,
  contentType: ContentType.Alert,
  targetLanguage: TargetLanguage.All,
  languageContent: [],
  targetUsers: [],
  targetGroups: [],
});

const createInitialState = (
  config: IAlertFormProviderConfig,
  services: IAlertFormServices
): IAlertFormState => ({
  newAlert: createInitialNewAlert(config.alertTypes),
  errors: {},
  isCreatingAlert: false,
  showPreview: true,
  showTemplates: true,
  createStep: "content",
  currentEntryMode: "templates" as const,
  creationProgress: [],
  lastCreateAttemptFailed: false,
  autoSaveDraftId: null,
  lastAutoSave: null,
  autoSaveStatus: "idle" as const,
  isAutoSaving: false,
  useMultiLanguage: false,
  supportedLanguages: [],
  languagePolicy: DEFAULT_LANGUAGE_POLICY,
  drafts: [],
  previousAlerts: [],
  copilotAvailability: "unknown",
  alertTypes: config.alertTypes,
  userTargetingEnabled: config.userTargetingEnabled,
  notificationsEnabled: config.notificationsEnabled,
  enableTargetSite: config.enableTargetSite ?? false,
  copilotEnabled: config.copilotEnabled ?? false,
});

// =============================================================================
// ACTIONS
// =============================================================================

export type AlertFormAction =
  | { type: "SET_FIELD"; field: keyof INewAlert; value: unknown }
  | { type: "SET_ALERT"; payload: Partial<INewAlert> }
  | { type: "SET_ERRORS"; errors: IFormErrors }
  | { type: "CLEAR_ERRORS"; fields?: string[] }
  | { type: "RESET_FORM" }
  | { type: "SET_IS_CREATING"; isCreating: boolean }
  | { type: "SET_SHOW_PREVIEW"; show: boolean }
  | { type: "SET_SHOW_TEMPLATES"; show: boolean }
  | { type: "SET_CREATE_STEP"; step: CreateWizardStep }
  | { type: "SET_ENTRY_MODE"; mode: "scratch" | "templates" | "drafts" | "previous" }
  | { type: "SET_CREATION_PROGRESS"; progress: ISiteValidationResult[] }
  | { type: "SET_LAST_CREATE_FAILED"; failed: boolean }
  | { type: "SET_AUTO_SAVE_STATUS"; status: "idle" | "pending" | "saving" | "saved" | "error" }
  | { type: "SET_AUTO_SAVE_DRAFT_ID"; id: string | null }
  | { type: "SET_LAST_AUTO_SAVE"; date: Date | null }
  | { type: "SET_IS_AUTO_SAVING"; isSaving: boolean }
  | { type: "SET_USE_MULTI_LANGUAGE"; useMultiLanguage: boolean }
  | { type: "SET_SUPPORTED_LANGUAGES"; languages: ISupportedLanguage[] }
  | { type: "SET_LANGUAGE_POLICY"; policy: ILanguagePolicy }
  | { type: "SET_DRAFTS"; drafts: IAlertItem[] }
  | { type: "SET_PREVIOUS_ALERTS"; alerts: IAlertItem[] }
  | { type: "SET_COPILOT_AVAILABILITY"; availability: "unknown" | "available" | "unavailable" }
  | { type: "UPDATE_ALERT_TYPES"; alertTypes: IAlertType[] }
  | { type: "LOAD_DRAFT"; draft: IAlertItem }
  | { type: "LOAD_PREVIOUS_ALERT"; alert: IAlertItem }
  | { type: "LOAD_TEMPLATE"; template: IAlertItem; useMultiLanguage: boolean };

// =============================================================================
// REDUCER
// =============================================================================

function alertFormReducer(state: IAlertFormState, action: AlertFormAction): IAlertFormState {
  switch (action.type) {
    case "SET_FIELD": {
      return {
        ...state,
        newAlert: {
          ...state.newAlert,
          [action.field]: action.value,
        },
      };
    }

    case "SET_ALERT": {
      return {
        ...state,
        newAlert: {
          ...state.newAlert,
          ...action.payload,
        },
      };
    }

    case "SET_ERRORS": {
      return {
        ...state,
        errors: { ...state.errors, ...action.errors },
      };
    }

    case "CLEAR_ERRORS": {
      if (!action.fields || action.fields.length === 0) {
        return { ...state, errors: {} };
      }
      const newErrors = { ...state.errors };
      action.fields.forEach((field) => {
        delete newErrors[field];
      });
      return { ...state, errors: newErrors };
    }

    case "RESET_FORM": {
      return {
        ...state,
        newAlert: createInitialNewAlert(state.alertTypes),
        errors: {},
        useMultiLanguage: false,
        createStep: "content",
        lastCreateAttemptFailed: false,
        currentEntryMode: "templates" as const,
        autoSaveDraftId: null,
        lastAutoSave: null,
        autoSaveStatus: "idle" as const,
        creationProgress: [],
      };
    }

    case "SET_IS_CREATING": {
      return { ...state, isCreatingAlert: action.isCreating };
    }

    case "SET_SHOW_PREVIEW": {
      return { ...state, showPreview: action.show };
    }

    case "SET_SHOW_TEMPLATES": {
      return { ...state, showTemplates: action.show };
    }

    case "SET_CREATE_STEP": {
      return { ...state, createStep: action.step };
    }

    case "SET_ENTRY_MODE": {
      return {
        ...state,
        currentEntryMode: action.mode,
        showTemplates: action.mode === "templates",
        showPreview: action.mode !== "templates" || state.showPreview,
      };
    }

    case "SET_CREATION_PROGRESS": {
      return { ...state, creationProgress: action.progress };
    }

    case "SET_LAST_CREATE_FAILED": {
      return { ...state, lastCreateAttemptFailed: action.failed };
    }

    case "SET_AUTO_SAVE_STATUS": {
      return { ...state, autoSaveStatus: action.status };
    }

    case "SET_AUTO_SAVE_DRAFT_ID": {
      return { ...state, autoSaveDraftId: action.id };
    }

    case "SET_LAST_AUTO_SAVE": {
      return { ...state, lastAutoSave: action.date };
    }

    case "SET_IS_AUTO_SAVING": {
      return { ...state, isAutoSaving: action.isSaving };
    }

    case "SET_USE_MULTI_LANGUAGE": {
      return { ...state, useMultiLanguage: action.useMultiLanguage };
    }

    case "SET_SUPPORTED_LANGUAGES": {
      return { ...state, supportedLanguages: action.languages };
    }

    case "SET_LANGUAGE_POLICY": {
      return { ...state, languagePolicy: action.policy };
    }

    case "SET_DRAFTS": {
      return { ...state, drafts: action.drafts };
    }

    case "SET_PREVIOUS_ALERTS": {
      return { ...state, previousAlerts: action.alerts };
    }

    case "SET_COPILOT_AVAILABILITY": {
      return {
        ...state,
        copilotAvailability: action.availability,
      };
    }

    case "UPDATE_ALERT_TYPES": {
      return {
        ...state,
        alertTypes: action.alertTypes,
        newAlert: {
          ...state.newAlert,
          AlertType:
            state.newAlert.AlertType &&
            action.alertTypes.find((t) => t.name === state.newAlert.AlertType)
              ? state.newAlert.AlertType
              : action.alertTypes.length > 0
                ? action.alertTypes[0].name
                : "",
        },
      };
    }

    case "LOAD_DRAFT": {
      const { draft } = action;
      return {
        ...state,
        newAlert: {
          title: draft.title,
          description: draft.description,
          AlertType: draft.AlertType,
          priority: draft.priority,
          isPinned: draft.isPinned,
          notificationType: draft.notificationType,
          linkUrl: draft.linkUrl || "",
          linkDescription: draft.linkDescription || "",
          targetSites: draft.targetSites || [],
          scheduledStart: draft.scheduledStart ? new Date(draft.scheduledStart) : undefined,
          scheduledEnd: draft.scheduledEnd ? new Date(draft.scheduledEnd) : undefined,
          contentType: ContentType.Alert,
          targetLanguage: draft.targetLanguage,
          languageContent: [],
          languageGroup: draft.languageGroup,
          targetUsers: draft.targetUsers || [],
          targetGroups: [],
        },
        currentEntryMode: "scratch",
        createStep: "content",
        autoSaveStatus: "pending",
        errors: {},
      };
    }

    case "LOAD_PREVIOUS_ALERT": {
      const { alert: sourceAlert } = action;
      return {
        ...state,
        newAlert: {
          ...state.newAlert,
          title: sourceAlert.title,
          description: sourceAlert.description,
          AlertType: sourceAlert.AlertType,
          priority: sourceAlert.priority,
          isPinned: sourceAlert.isPinned,
          notificationType: sourceAlert.notificationType,
          linkUrl: sourceAlert.linkUrl || "",
          linkDescription: sourceAlert.linkDescription || "",
          targetSites: sourceAlert.targetSites || state.newAlert.targetSites,
          scheduledStart: sourceAlert.scheduledStart
            ? new Date(sourceAlert.scheduledStart)
            : undefined,
          scheduledEnd: sourceAlert.scheduledEnd
            ? new Date(sourceAlert.scheduledEnd)
            : undefined,
          contentType: ContentType.Alert,
          targetLanguage: sourceAlert.targetLanguage || state.newAlert.targetLanguage,
          languageContent: [],
          languageGroup: undefined,
          targetUsers: [],
          targetGroups: [],
        },
        errors: {},
        currentEntryMode: "scratch",
        createStep: "content",
        autoSaveStatus: "pending",
      };
    }

    case "LOAD_TEMPLATE": {
      const { template, useMultiLanguage: templateMultiLanguage } = action;
      return {
        ...state,
        newAlert: {
          ...state.newAlert,
          title: template.title,
          description: template.description,
          AlertType: template.AlertType,
          priority: template.priority,
          notificationType: template.notificationType,
          isPinned: template.isPinned,
          linkUrl: template.linkUrl || "",
          linkDescription: template.linkDescription || "",
          contentType: ContentType.Alert,
          languageContent: templateMultiLanguage && template.languageGroup ? [] : [],
          languageGroup: templateMultiLanguage && template.languageGroup
            ? template.languageGroup
            : undefined,
        },
        currentEntryMode: "scratch",
        createStep: "content",
        errors: {},
      };
    }

    default:
      return state;
  }
}

// =============================================================================
// CONTEXT
// =============================================================================

interface IAlertFormContextValue {
  state: IAlertFormState;
  dispatch: React.Dispatch<AlertFormAction>;
}

const AlertFormContext = React.createContext<IAlertFormContextValue | undefined>(undefined);
const AlertFormServicesContext = React.createContext<IAlertFormServices | undefined>(undefined);

// =============================================================================
// PROVIDER CONFIG
// =============================================================================

export interface IAlertFormProviderConfig {
  alertTypes: IAlertType[];
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite?: boolean;
  copilotEnabled?: boolean;
  languageUpdateTrigger?: number;
}

export interface IAlertFormServices {
  siteDetector: SiteContextDetector;
  alertService: SharePointAlertService;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  languageService: LanguageAwarenessService;
}

export interface IAlertFormProviderProps {
  config: IAlertFormProviderConfig;
  services: IAlertFormServices;
  children: React.ReactNode;
  onDirtyStateChange?: (hasUnsavedChanges: boolean) => void;
}

// =============================================================================
// PROVIDER COMPONENT
// =============================================================================

export const AlertFormProvider: React.FC<IAlertFormProviderProps> = ({
  config,
  services,
  children,
  onDirtyStateChange,
}) => {
  const [state, dispatch] = React.useReducer(
    alertFormReducer,
    createInitialState(config, services)
  );

  // Update alert types when they change externally
  React.useEffect(() => {
    dispatch({ type: "UPDATE_ALERT_TYPES", alertTypes: config.alertTypes });
  }, [config.alertTypes]);

  // Track dirty state
  React.useEffect(() => {
    if (!onDirtyStateChange) {
      return;
    }

    const hasLanguageContent = state.newAlert.languageContent.some(
      (item) =>
        item.title.trim().length > 0 ||
        item.description.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().length > 0 ||
        (item.linkDescription || "").trim().length > 0
    );

    const hasUnsavedChanges =
      state.currentEntryMode === "scratch" &&
      (state.newAlert.title.trim().length > 0 ||
        state.newAlert.description.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().length > 0 ||
        state.newAlert.linkUrl.trim().length > 0 ||
        state.newAlert.linkDescription.trim().length > 0 ||
        hasLanguageContent);

    onDirtyStateChange(hasUnsavedChanges);
  }, [
    state.newAlert,
    state.currentEntryMode,
    onDirtyStateChange,
  ]);

  // Memoize context value to prevent unnecessary re-renders
  const contextValue = React.useMemo(
    () => ({ state, dispatch }),
    [state, dispatch]
  );

  // Memoize services to prevent unnecessary re-renders
  const servicesValue = React.useMemo(
    () => services,
    [services.siteDetector, services.alertService, services.graphClient, services.context, services.languageService]
  );

  return (
    <AlertFormServicesContext.Provider value={servicesValue}>
      <AlertFormContext.Provider value={contextValue}>
        {children}
      </AlertFormContext.Provider>
    </AlertFormServicesContext.Provider>
  );
};

// =============================================================================
// HOOKS
// =============================================================================

/**
 * Hook to access the alert form state and dispatch
 * @throws Error if used outside of AlertFormProvider
 */
export function useAlertForm(): IAlertFormContextValue {
  const context = React.useContext(AlertFormContext);
  if (context === undefined) {
    throw new Error("useAlertForm must be used within an AlertFormProvider");
  }
  return context;
}

/**
 * Hook to access only the alert form state
 * Useful when you only need to read values
 */
export function useAlertFormState(): IAlertFormState {
  const { state } = useAlertForm();
  return state;
}

/**
 * Hook to access only the alert form dispatch
 * Useful when you only need to dispatch actions
 */
export function useAlertFormDispatch(): React.Dispatch<AlertFormAction> {
  const { dispatch } = useAlertForm();
  return dispatch;
}

/**
 * Hook for field-level state access and updates
 * Returns [value, setValue] tuple similar to useState
 */
export function useAlertFormField<K extends keyof INewAlert>(
  field: K
): [INewAlert[K], (value: INewAlert[K]) => void] {
  const { state, dispatch } = useAlertForm();

  const setValue = React.useCallback(
    (value: INewAlert[K]) => {
      dispatch({ type: "SET_FIELD", field, value });
    },
    [dispatch, field]
  );

  return [state.newAlert[field], setValue];
}

/**
 * Hook to access the alert form services
 * @throws Error if used outside of AlertFormProvider
 */
export function useAlertFormServices(): IAlertFormServices {
  const context = React.useContext(AlertFormServicesContext);
  if (context === undefined) {
    throw new Error("useAlertFormServices must be used within an AlertFormProvider");
  }
  return context;
}

/**
 * Hook for computed/derived state values
 * Memoizes expensive computations
 */
export function useAlertFormComputed<T>(
  compute: (state: IAlertFormState) => T,
  deps: React.DependencyList = []
): T {
  const { state } = useAlertForm();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  return React.useMemo(() => compute(state), [state, ...deps]);
}

export default AlertFormContext;
