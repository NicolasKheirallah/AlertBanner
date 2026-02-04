import * as React from "react";
import { Save24Regular, Eye24Regular, Dismiss24Regular, Drafts24Regular, DocumentArrowLeft24Regular } from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointSelect,
  SharePointToggle,
  SharePointSection,
  ISharePointSelectOption,
  SharePointPeoplePicker
} from "../../UI/SharePointControls";
import SharePointRichTextEditor from "../../UI/SharePointRichTextEditor";
import AlertPreview from "../../UI/AlertPreview";
import AlertTemplates, { IAlertTemplate } from "../../UI/AlertTemplates";
import SiteSelector from "../../UI/SiteSelector";
import MultiLanguageContentEditor from "../../UI/MultiLanguageContentEditor";
import { AlertPriority, NotificationType, IAlertType, ContentType, TargetLanguage, IPersonField, IAlertItem, TranslationStatus } from "../../Alerts/IAlerts";
import { LanguageAwarenessService, ILanguageContent, ISupportedLanguage } from "../../Services/LanguageAwarenessService";
import { DEFAULT_LANGUAGE_POLICY, ILanguagePolicy } from "../../Services/LanguagePolicyService";
import { SiteContextDetector, ISiteValidationResult } from "../../Utils/SiteContextDetector";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { logger } from '../../Services/LoggerService';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { NotificationService } from '../../Services/NotificationService';
import styles from "../AlertSettings.module.scss";
import { validateAlertData, IFormErrors as IValidationErrors } from "../../Utils/AlertValidation";
import { useLanguageOptions } from "../../Hooks/useLanguageOptions";
import { usePriorityOptions } from "../../Hooks/usePriorityOptions";
import { DateUtils } from "../../Utils/DateUtils";
import { StringUtils } from "../../Utils/StringUtils";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

const STRINGS_DICTIONARY = strings as unknown as Record<string, string>;

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
  languageGroup?: string; // For multi-language alerts to share resources
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
}

export interface IFormErrors {
  title?: string;
  description?: string;
  AlertType?: string;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  // Index signature for dynamic language error keys
  [key: string]: string | undefined;
}

export interface ICreateAlertTabProps {
  newAlert: INewAlert;
  setNewAlert: React.Dispatch<React.SetStateAction<INewAlert>>;
  errors: IFormErrors;
  setErrors: React.Dispatch<React.SetStateAction<IFormErrors>>;
  alertTypes: IAlertType[];
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite?: boolean; // Optional - defaults to false
  siteDetector: SiteContextDetector;
  alertService: SharePointAlertService;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  creationProgress: ISiteValidationResult[];
  setCreationProgress: React.Dispatch<React.SetStateAction<ISiteValidationResult[]>>;
  isCreatingAlert: boolean;
  setIsCreatingAlert: React.Dispatch<React.SetStateAction<boolean>>;
  showPreview: boolean;
  setShowPreview: React.Dispatch<React.SetStateAction<boolean>>;
  showTemplates: boolean;
  setShowTemplates: React.Dispatch<React.SetStateAction<boolean>>;
  languageUpdateTrigger?: number;
}

const CreateAlertTab: React.FC<ICreateAlertTabProps> = ({
  newAlert,
  setNewAlert,
  errors,
  setErrors,
  alertTypes,
  userTargetingEnabled,
  notificationsEnabled,
  enableTargetSite = false, // Default to false
  siteDetector,
  alertService,
  graphClient,
  context,
  creationProgress,
  setCreationProgress,
  isCreatingAlert,
  setIsCreatingAlert,
  showPreview,
  setShowPreview,
  showTemplates,
  setShowTemplates,
  languageUpdateTrigger
}) => {
  // Priority options - using shared hook
  const priorityOptions = usePriorityOptions();

  // Notification type options with detailed descriptions
  const notificationOptions: ISharePointSelectOption[] = React.useMemo(() => ([
    { value: NotificationType.None, label: strings.CreateAlertNotificationNoneDescription },
    { value: NotificationType.Browser, label: strings.CreateAlertNotificationBrowserDescription },
    { value: NotificationType.Email, label: strings.CreateAlertNotificationEmailDescription },
    { value: NotificationType.Both, label: strings.CreateAlertNotificationBothDescription }
  ]), []);

  // Alert type options
  const alertTypeOptions: ISharePointSelectOption[] = alertTypes.map(type => ({
    value: type.name,
    label: type.name
  }));

  // Content type options
  const contentTypeOptions: ISharePointSelectOption[] = React.useMemo(() => ([
    { value: ContentType.Alert, label: strings.CreateAlertContentTypeAlertDescription },
    { value: ContentType.Template, label: strings.CreateAlertContentTypeTemplateDescription }
  ]), []);

  // Language awareness state
  const [languageService] = React.useState(() => new LanguageAwarenessService(graphClient, context));
  const [supportedLanguages, setSupportedLanguages] = React.useState<ISupportedLanguage[]>([]);
  const [useMultiLanguage, setUseMultiLanguage] = React.useState(false);
  const [languagePolicy, setLanguagePolicy] = React.useState<ILanguagePolicy>(DEFAULT_LANGUAGE_POLICY);
  const notificationService = React.useMemo(() => NotificationService.getInstance(context), [context]);

  // Draft state
  const [drafts, setDrafts] = React.useState<IAlertItem[]>([]);
  const [showDrafts, setShowDrafts] = React.useState(false);
  const [autoSaveDraftId, setAutoSaveDraftId] = React.useState<string | null>(null);
  const [lastAutoSave, setLastAutoSave] = React.useState<Date | null>(null);
  const [isAutoSaving, setIsAutoSaving] = React.useState(false);

  // Language targeting options - using shared hook
  const languageOptions = useLanguageOptions(supportedLanguages);

  // Load supported languages from SharePoint (actual enabled ones)
  const loadSupportedLanguages = React.useCallback(async () => {
    try {
      // Get the base language definitions
      const baseLanguages = LanguageAwarenessService.getSupportedLanguages();
      
      // Get the actually supported languages from SharePoint columns
      const supportedLanguageCodes = await alertService.getSupportedLanguages();
      
      logger.info('CreateAlertTab', `Available language columns: ${supportedLanguageCodes.length}`);
      
      // Update the base languages with the actual status
      const updatedLanguages = baseLanguages.map(lang => ({
        ...lang,
        columnExists: supportedLanguageCodes.includes(lang.code) || lang.code === TargetLanguage.EnglishUS,
        isSupported: supportedLanguageCodes.includes(lang.code) || lang.code === TargetLanguage.EnglishUS
      }));
      
      setSupportedLanguages(updatedLanguages);
      logger.debug('CreateAlertTab', 'Updated supported languages', { supportedLanguages: updatedLanguages.filter(l => l.isSupported).map(l => l.code) });
    } catch (error) {
      logger.error('CreateAlertTab', 'Error loading supported languages', error);
      // Fallback to default with English only
      const defaultLanguages = LanguageAwarenessService.getSupportedLanguages();
      setSupportedLanguages(defaultLanguages.map(lang => ({
        ...lang,
        isSupported: lang.code === TargetLanguage.EnglishUS,
        columnExists: lang.code === TargetLanguage.EnglishUS
      })));
    }
  }, [alertService]);

  React.useEffect(() => {
    loadSupportedLanguages();
  }, [loadSupportedLanguages, languageUpdateTrigger]);

  React.useEffect(() => {
    let isMounted = true;
    alertService.getLanguagePolicy()
      .then(policy => {
        if (isMounted) {
          setLanguagePolicy(policy);
        }
      })
      .catch(error => {
        logger.warn('CreateAlertTab', 'Failed to load language policy', error);
      });
    return () => { isMounted = false; };
  }, [alertService]);

  // Load drafts
  const loadDrafts = React.useCallback(async () => {
    try {
      const siteId = alertService.getCurrentSiteId();
      const userDrafts = await alertService.getDraftAlerts(siteId);
      setDrafts(userDrafts);
    } catch (error) {
      logger.warn('CreateAlertTab', 'Could not load drafts', error);
      setDrafts([]);
    }
  }, [alertService]);

  React.useEffect(() => {
    loadDrafts();
  }, [loadDrafts]);

  // Load from draft
  const handleLoadDraft = React.useCallback((draft: IAlertItem) => {
    setNewAlert({
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
      contentType: ContentType.Alert, // Convert draft to alert
      targetLanguage: draft.targetLanguage,
      languageContent: [],
      languageGroup: draft.languageGroup
    });
    setShowDrafts(false);
    setShowTemplates(false);
  }, [setNewAlert, setShowTemplates]);

  // Initialize language content when multi-language is enabled
  React.useEffect(() => {
    if (useMultiLanguage && newAlert.languageContent.length === 0) {
      // Start with English by default and generate a languageGroup ID
      setNewAlert(prev => {
        const englishContent: ILanguageContent = {
          language: TargetLanguage.EnglishUS,
          title: prev.title,
          description: prev.description,
          linkDescription: prev.linkUrl ? prev.linkDescription : undefined,
          translationStatus: languagePolicy.workflow.enabled ? languagePolicy.workflow.defaultStatus : TranslationStatus.Approved
        };
        // Generate a unique languageGroup ID for sharing resources (images) across language variants
        const languageGroup = prev.languageGroup || `lg-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
        return { ...prev, languageContent: [englishContent], languageGroup };
      });
    } else if (!useMultiLanguage && newAlert.languageContent.length > 0) {
      // When switching back to single language, use the first language's content
      setNewAlert(prev => {
        const firstLang = prev.languageContent[0];
        if (firstLang) {
          return {
            ...prev,
            title: firstLang.title,
            description: firstLang.description,
            linkDescription: firstLang.linkDescription || '',
            languageContent: [],
            languageGroup: undefined // Clear languageGroup for single language
          };
        }
        return prev;
      });
    }
  }, [useMultiLanguage, newAlert.languageContent.length]);

  const handleTemplateSelect = React.useCallback(async (template: IAlertTemplate) => {
    try {
      // Get all templates from current site to include all fields
      const templateAlerts = await alertService.getTemplateAlerts(alertService.getCurrentSiteId());
      const fullTemplate = templateAlerts.find(t => t.id === template.id);

      if (fullTemplate) {
        // Use the full template data
        setNewAlert(prev => ({
          ...prev,
          title: fullTemplate.title,
          description: fullTemplate.description,
          AlertType: fullTemplate.AlertType,
          priority: fullTemplate.priority,
          notificationType: fullTemplate.notificationType,
          isPinned: fullTemplate.isPinned,
          linkUrl: fullTemplate.linkUrl || "",
          linkDescription: fullTemplate.linkDescription || "",
          contentType: ContentType.Alert, // New alert should be Alert, not Template
        }));

        // If multi-language is enabled, load language variants of the template
        if (useMultiLanguage && fullTemplate.languageGroup) {
          try {
            // Filter the already-loaded templates for language variants
            const languageVariants = templateAlerts.filter(a =>
              a.languageGroup === fullTemplate.languageGroup
            );

            // Convert to language content format
            const languageContent: ILanguageContent[] = languageVariants.map(variant => ({
              language: variant.targetLanguage,
              title: variant.title,
              description: variant.description,
              linkDescription: variant.linkDescription || ""
            }));

            if (languageContent.length > 0) {
              setNewAlert(prev => ({
                ...prev,
                languageContent,
                targetLanguage: languageContent[0].language // Set primary language
              }));
            }
          } catch (error) {
            logger.warn('CreateAlertTab', 'Could not load template language variants', error);
          }
        }
      } else {
        // Fallback to basic template data
        setNewAlert(prev => ({
          ...prev,
          title: template.template.title,
          description: template.template.description,
          priority: template.template.priority,
          notificationType: template.template.notificationType,
          isPinned: template.template.isPinned,
          linkUrl: template.template.linkUrl || "",
          linkDescription: template.template.linkDescription || "",
          contentType: ContentType.Alert, // New alert should be Alert, not Template
        }));
      }
    } catch (error) {
      logger.error('CreateAlertTab', 'Failed to load template', error);
      // Fallback to basic template data
      setNewAlert(prev => ({
        ...prev,
        title: template.template.title,
        description: template.template.description,
        priority: template.template.priority,
        notificationType: template.template.notificationType,
        isPinned: template.template.isPinned,
        linkUrl: template.template.linkUrl || "",
        linkDescription: template.template.linkDescription || "",
        contentType: ContentType.Alert, // New alert should be Alert, not Template
      }));
    }
    
    setShowTemplates(false);
  }, [alertService, setNewAlert, setShowTemplates, useMultiLanguage]);

  // Validation using shared utility
  const validateForm = React.useCallback((): boolean => {
    const validationErrors = validateAlertData(newAlert, {
      useMultiLanguage,
      languagePolicy,
      tenantDefaultLanguage: languageService.getTenantDefaultLanguage(),
      getString: (key: string, ...args: Array<string | number>) => {
        const template = STRINGS_DICTIONARY[key] ?? key;
        if (args.length === 0) {
          return template;
        }
        const formattedArgs = args.map(arg => arg.toString());
        return Text.format(template, ...formattedArgs);
      }
    });

    setErrors(validationErrors);
    return Object.keys(validationErrors).length === 0;
  }, [newAlert, useMultiLanguage, languagePolicy, languageService, setErrors]);
  
  const handlePeoplePickerChange = React.useCallback((items: any[]) => {
    const users: IPersonField[] = [];
    const groups: IPersonField[] = [];
    
    items.forEach(item => {
      const personField: IPersonField = {
        id: item.id,
        displayName: item.text,
        email: item.secondaryText, // Usually email
        loginName: item.loginName,
        isGroup: item.id ? item.id.indexOf('c:0(.s|true') === -1 : false // Simple check, might need refinement based on PrincipalType
      };
      
      // PnP PeoplePicker returns items. We need to distinguish users and groups.
      // PrincipalType: User = 1, DistributionList = 2, SecurityGroup = 4, SharePointGroup = 8
      // But the items array might not have the numeric type directly accessible in the same way.
      // Let's rely on the image or type if available, but for now we'll assume:
      // If it has an email, it's likely a user. If not, it might be a group.
      // Better approach: Check item.imageInitials or item.imageUrl
      
      if (item.imageInitials || (item.secondaryText && item.secondaryText.indexOf('@') > -1)) {
        personField.isGroup = false;
        users.push(personField);
      } else {
        personField.isGroup = true;
        groups.push(personField);
      }
    });
    
    setNewAlert(prev => ({
      ...prev,
      targetUsers: users,
      targetGroups: groups
    }));
  }, [setNewAlert]);

  const handleCreateAlert = React.useCallback(async () => {
    if (!validateForm()) return;

    setIsCreatingAlert(true);
    setCreationProgress([]);

    try {
      if (useMultiLanguage && newAlert.languageContent.length > 0) {
        // Create multi-language alerts
        const multiLanguageAlert = languageService.createMultiLanguageAlert({
          AlertType: newAlert.AlertType,
          priority: newAlert.priority,
          isPinned: newAlert.isPinned,
          linkUrl: newAlert.linkUrl,
          notificationType: newAlert.notificationType,
          createdDate: new Date().toISOString(),
          createdBy: context.pageContext.user.displayName,
          contentType: newAlert.contentType,
          targetLanguage: TargetLanguage.All,
          status: 'Active' as 'Active' | 'Expired' | 'Scheduled',
          targetSites: newAlert.targetSites,
          id: '0',
          targetUsers: [...(newAlert.targetUsers || []), ...(newAlert.targetGroups || [])],
          languageGroup: newAlert.languageGroup
        }, newAlert.languageContent);

        const alertItems = languageService.generateAlertItems(multiLanguageAlert);
        
        // Create each language variant
        for (const alertItem of alertItems) {
          const alertData = {
            ...alertItem,
            targetSites: newAlert.targetSites,
            scheduledStart: newAlert.scheduledStart?.toISOString(),
            scheduledEnd: newAlert.scheduledEnd?.toISOString()
          };
          await alertService.createAlert(alertData);
        }
        
        setCreationProgress([{
          siteId: "success",
          siteName: Text.format(strings.CreateAlertMultiLanguageSuccessStatus, alertItems.length),
          hasAccess: true,
          canCreateAlerts: true,
          permissionLevel: "success",
          error: ""
        }]);
        notificationService.showSuccess(Text.format(strings.CreateAlertMultiLanguageSuccessStatus, alertItems.length), 'Alerts Created');
      } else {
        // Create single language alert
        const alertData = {
          title: newAlert.title,
          description: newAlert.description,
          AlertType: newAlert.AlertType,
          priority: newAlert.priority,
          isPinned: newAlert.isPinned,
          notificationType: newAlert.notificationType,
          linkUrl: newAlert.linkUrl,
          linkDescription: newAlert.linkDescription,
          targetSites: newAlert.targetSites,
          scheduledStart: newAlert.scheduledStart?.toISOString(),
          scheduledEnd: newAlert.scheduledEnd?.toISOString(),
          contentType: newAlert.contentType,
          targetLanguage: newAlert.targetLanguage,
          targetUsers: [...(newAlert.targetUsers || []), ...(newAlert.targetGroups || [])],
          translationStatus: languagePolicy.workflow.enabled ? languagePolicy.workflow.defaultStatus : TranslationStatus.Approved
        };

        await alertService.createAlert(alertData);
        
        setCreationProgress([{
          siteId: "success",
          siteName: strings.CreateAlertCreationSuccessStatus,
          hasAccess: true,
          canCreateAlerts: true,
          permissionLevel: "success",
          error: ""
        }]);
        notificationService.showSuccess(strings.CreateAlertCreationSuccessStatus, 'Alert Created');
      }

      // Reset form on success
      setNewAlert({
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
        targetGroups: []
      });
      setUseMultiLanguage(false);
      setShowTemplates(true);
    } catch (error) {
      logger.error('CreateAlertTab', 'Error creating alert', error);
      setCreationProgress([{
        siteId: "error",
        siteName: strings.CreateAlertCreationErrorStatus,
        hasAccess: false,
        canCreateAlerts: false,
        permissionLevel: "error",
        error: error instanceof Error ? error.message : strings.CreateAlertUnknownError
      }]);
      notificationService.showError(strings.CreateAlertCreationErrorStatus, 'Creation Failed');
    } finally {
      setIsCreatingAlert(false);
    }
  }, [validateForm, setIsCreatingAlert, setCreationProgress, alertService, newAlert, setNewAlert, alertTypes, setShowTemplates, useMultiLanguage, languageService, context.pageContext.user.displayName, languagePolicy]);

  const handleSaveAsDraft = React.useCallback(async () => {
    // Basic validation for drafts (less strict than full validation)
    if (!newAlert.title || newAlert.title.trim().length < 3) {
      setErrors({ title: "Draft must have a title (at least 3 characters)" });
      return;
    }

    // Ensure AlertType is set (required field even for drafts)
    if (!newAlert.AlertType || !alertTypes.find(t => t.name === newAlert.AlertType)) {
      setErrors({ AlertType: "Please select an alert type before saving draft" });
      return;
    }

    setIsCreatingAlert(true);
    try {
      const draftData: Partial<IAlertItem> = {
        title: newAlert.title,
        description: newAlert.description,
        AlertType: newAlert.AlertType,
        priority: newAlert.priority,
        isPinned: newAlert.isPinned,
        notificationType: newAlert.notificationType,
        linkUrl: newAlert.linkUrl,
        linkDescription: newAlert.linkDescription,
        targetSites: newAlert.targetSites,
        scheduledStart: newAlert.scheduledStart?.toISOString(),
        scheduledEnd: newAlert.scheduledEnd?.toISOString(),
        contentType: ContentType.Draft,
        targetLanguage: newAlert.targetLanguage,
        languageGroup: newAlert.languageGroup,
        createdDate: new Date().toISOString(),
        createdBy: context.pageContext.user.displayName
      };

      await alertService.saveDraft(draftData);

      setCreationProgress([{
        siteId: "success",
        siteName: "Draft Saved Successfully",
        hasAccess: true,
        canCreateAlerts: true,
        permissionLevel: "success",
        error: ""
      }]);

      // Reload drafts and reset form after saving
      await loadDrafts();
      setTimeout(() => {
        resetForm();
        setCreationProgress([]);
      }, 2000);
    } catch (error) {
      logger.error('CreateAlertTab', 'Error saving draft', error);
      setCreationProgress([{
        siteId: "error",
        siteName: "Draft Save Error",
        hasAccess: false,
        canCreateAlerts: false,
        permissionLevel: "error",
        error: error instanceof Error ? error.message : "Unknown error occurred"
      }]);
    } finally {
      setIsCreatingAlert(false);
    }
  }, [newAlert, alertService, context.pageContext.user.displayName, setErrors, setIsCreatingAlert, setCreationProgress, loadDrafts, alertTypes]);

  // Auto-save draft functionality
  const autoSaveDraft = React.useCallback(async () => {
    // Only auto-save if there's meaningful content and AlertType is selected
    if (!newAlert.title || newAlert.title.trim().length < 3 || !newAlert.AlertType) {
      return;
    }

    // Don't auto-save if currently creating/saving
    if (isCreatingAlert || isAutoSaving) {
      return;
    }

    setIsAutoSaving(true);
    try {
      const draftData: Partial<IAlertItem> = {
        id: autoSaveDraftId || undefined,
        title: `[Auto-saved] ${newAlert.title}`,
        description: newAlert.description,
        AlertType: newAlert.AlertType,
        priority: newAlert.priority,
        isPinned: newAlert.isPinned,
        notificationType: newAlert.notificationType,
        linkUrl: newAlert.linkUrl,
        linkDescription: newAlert.linkDescription,
        targetSites: newAlert.targetSites,
        scheduledStart: newAlert.scheduledStart?.toISOString(),
        scheduledEnd: newAlert.scheduledEnd?.toISOString(),
        contentType: ContentType.Draft,
        targetLanguage: newAlert.targetLanguage,
        languageGroup: newAlert.languageGroup,
        createdDate: new Date().toISOString(),
        createdBy: context.pageContext.user.displayName
      };

      const savedDraft = await alertService.saveDraft(draftData);
      setAutoSaveDraftId(savedDraft.id);
      setLastAutoSave(new Date());
      logger.debug('CreateAlertTab', 'Auto-saved draft', { id: savedDraft.id, title: savedDraft.title });
    } catch (error) {
      logger.warn('CreateAlertTab', 'Auto-save failed', error);
      // Don't show error to user for auto-save failures
    } finally {
      setIsAutoSaving(false);
    }
  }, [newAlert, alertService, context.pageContext.user.displayName, isCreatingAlert, isAutoSaving, autoSaveDraftId]);

  // Auto-save effect - runs every 30 seconds
  React.useEffect(() => {
    const autoSaveInterval = setInterval(() => {
      autoSaveDraft();
    }, 30000); // 30 seconds

    return () => clearInterval(autoSaveInterval);
  }, [autoSaveDraft]);

  const resetForm = React.useCallback(() => {
    setNewAlert({
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
      targetGroups: []
    });
    setErrors({});
    setUseMultiLanguage(false);
    setShowTemplates(true);
    setAutoSaveDraftId(null);
    setLastAutoSave(null);
  }, [setNewAlert, alertTypes, setErrors, setShowTemplates]);

  const getCurrentAlertType = React.useCallback((): IAlertType | undefined => {
    return alertTypes.find(type => type.name === newAlert.AlertType);
  }, [alertTypes, newAlert.AlertType]);

  return (
    <div className={styles.tabContent}>
      {(showTemplates || showDrafts) && (
        <div className={styles.templatesSection}>
          <div className={styles.templateActions}>
            <SharePointButton
              variant={showTemplates && !showDrafts ? "primary" : "secondary"}
              onClick={() => { setShowTemplates(true); setShowDrafts(false); }}
            >
              {strings.CreateAlertTemplatesButtonLabel}
            </SharePointButton>
            <SharePointButton
              variant={showDrafts ? "primary" : "secondary"}
              onClick={() => { setShowTemplates(false); setShowDrafts(true); }}
            >
              {Text.format(strings.CreateAlertDraftsButtonLabel, drafts.length)}
            </SharePointButton>
          </div>

          {showTemplates && !showDrafts && (
            <AlertTemplates
              onSelectTemplate={handleTemplateSelect}
              graphClient={graphClient}
              context={context}
              alertService={alertService}
              className={styles.templates}
            />
          )}

          {showDrafts && (
            <div className={styles.alertsList}>
              {drafts.length === 0 ? (
                <div className={styles.emptyState}>
                  <p>{strings.CreateAlertDraftsEmptyMessage}</p>
                </div>
              ) : (
                drafts.map(draft => (
                  <div key={draft.id} className={styles.alertCard}>
                    <div className={styles.alertCardHeader}>
                      <h4>{draft.title}</h4>
                      <small>
                        {DateUtils.formatForDisplay(draft.createdDate)}
                      </small>
                    </div>
                    <div className={styles.alertCardContent}>
                      <div dangerouslySetInnerHTML={{ __html: StringUtils.truncate(draft.description, 100) }} />
                    </div>
                    <div className={styles.alertCardActions}>
                      <SharePointButton
                        variant="primary"
                        onClick={() => handleLoadDraft(draft)}
                        icon={<DocumentArrowLeft24Regular />}
                      >
                        {strings.CreateAlertLoadDraftButton}
                      </SharePointButton>
                    </div>
                  </div>
                ))
              )}
            </div>
          )}

          <div className={styles.templateActions}>
            <SharePointButton
              variant="secondary"
              onClick={() => { setShowTemplates(false); setShowDrafts(false); }}
            >
              {strings.CreateAlertStartFromScratchButton}
            </SharePointButton>
          </div>
        </div>
      )}

      {!showTemplates && !showDrafts && (
        <div className={styles.alertForm}>
          <div className={styles.formWithPreview}>
            <div className={styles.formColumn}>
              <SharePointSection title={strings.CreateAlertSectionContentClassificationTitle}>
                <SharePointSelect
                  label={strings.ContentTypeLabel}
                  value={newAlert.contentType}
                  onChange={(value) => setNewAlert(prev => ({ ...prev, contentType: value as ContentType }))}
                  options={contentTypeOptions}
                  required
                  description={strings.CreateAlertSectionContentClassificationDescription}
                />

                <div className={styles.languageModeSelector}>
                  <label className={styles.fieldLabel}>{strings.CreateAlertLanguageConfigurationLabel}</label>
                  <div className={styles.languageOptions}>
                    <SharePointButton
                      variant={!useMultiLanguage ? "primary" : "secondary"}
                      onClick={() => setUseMultiLanguage(false)}
                    >
                      {strings.CreateAlertSingleLanguageButton}
                    </SharePointButton>
                    <SharePointButton
                      variant={useMultiLanguage ? "primary" : "secondary"}
                      onClick={() => setUseMultiLanguage(true)}
                    >
                      {strings.CreateAlertMultiLanguageButton}
                    </SharePointButton>
                  </div>
                </div>
              </SharePointSection>

              {!useMultiLanguage ? (
                <>
                  <SharePointSection title={strings.CreateAlertSectionLanguageTargetingTitle}>
                    <SharePointSelect
                      label={strings.CreateAlertTargetLanguageLabel}
                      value={newAlert.targetLanguage}
                      onChange={(value) => setNewAlert(prev => ({ ...prev, targetLanguage: value as TargetLanguage }))}
                      options={languageOptions}
                      required
                      description={strings.CreateAlertSectionLanguageTargetingDescription}
                    />
                  </SharePointSection>
                  
                  {userTargetingEnabled && (
                    <SharePointSection title={strings.CreateAlertSectionUserTargetingTitle || "User Targeting"}>
                      <SharePointPeoplePicker
                        context={context}
                        titleText={strings.CreateAlertPeoplePickerLabel || "Target specific users or groups"}
                        personSelectionLimit={50}
                        groupName={""} // Leave empty to search all
                        showtooltip={true}
                        required={false}
                        disabled={isCreatingAlert}
                        onChange={handlePeoplePickerChange}
                        defaultSelectedUsers={[
                          ...(newAlert.targetUsers?.map(u => u.email || u.loginName || u.displayName) || []),
                          ...(newAlert.targetGroups?.map(g => g.loginName || g.displayName) || [])
                        ]}
                        description={strings.CreateAlertPeoplePickerDescription || "Leave empty to show to everyone. Select users or groups to restrict visibility."}
                      />
                    </SharePointSection>
                  )}

                  <SharePointSection title={strings.CreateAlertSectionBasicInformationTitle}>
                    <SharePointInput
                      label={strings.AlertTitle}
                      value={newAlert.title}
                      onChange={(value) => {
                        setNewAlert(prev => ({ ...prev, title: value }));
                        setErrors(prev => prev.title ? { ...prev, title: undefined } : prev);
                      }}
                      placeholder={strings.CreateAlertTitlePlaceholder}
                      required
                      error={errors.title}
                      description={strings.CreateAlertTitleDescription}
                    />

                    <SharePointRichTextEditor
                      label={strings.AlertDescription}
                      value={newAlert.description}
                      onChange={(value) => {
                        setNewAlert(prev => ({ ...prev, description: value }));
                        if (errors.description) setErrors(prev => ({ ...prev, description: undefined }));
                      }}
                      context={context}
                      placeholder={strings.CreateAlertDescriptionPlaceholder}
                      required
                      error={errors.description}
                      description={strings.CreateAlertDescriptionHelp}
                      imageFolderName={newAlert.languageGroup || newAlert.title || 'Untitled_Alert'}
                    />
                  </SharePointSection>
                </>
              ) : (
                <SharePointSection title={strings.MultiLanguageContent}>
                  <MultiLanguageContentEditor
                    content={newAlert.languageContent}
                    onContentChange={(content) => setNewAlert(prev => ({ ...prev, languageContent: content }))}
                    availableLanguages={supportedLanguages}
                    errors={errors}
                    linkUrl={newAlert.linkUrl}
                    context={context}
                    imageFolderName={newAlert.languageGroup}
                    tenantDefaultLanguage={languageService.getTenantDefaultLanguage()}
                    languagePolicy={languagePolicy}
                  />
                </SharePointSection>
              )}

              <SharePointSection title={strings.CreateAlertConfigurationSectionTitle}>
                <SharePointSelect
                  label={strings.AlertType}
                  value={newAlert.AlertType}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, AlertType: value }));
                    if (errors.AlertType) setErrors(prev => ({ ...prev, AlertType: undefined }));
                  }}
                  options={alertTypeOptions}
                  required
                  error={errors.AlertType}
                  description={strings.CreateAlertConfigurationDescription}
                />

                <SharePointSelect
                  label={strings.CreateAlertPriorityLabel}
                  value={newAlert.priority}
                  onChange={(value) => setNewAlert(prev => ({ ...prev, priority: value as AlertPriority }))}
                  options={priorityOptions}
                  required
                  description={strings.CreateAlertPriorityDescription}
                />

                <SharePointToggle
                  label={strings.CreateAlertPinLabel}
                  checked={newAlert.isPinned}
                  onChange={(checked) => setNewAlert(prev => ({ ...prev, isPinned: checked }))}
                  description={strings.CreateAlertPinDescription}
                />

                {notificationsEnabled && (
                  <SharePointSelect
                    label={strings.CreateAlertNotificationLabel}
                    value={newAlert.notificationType}
                    onChange={(value) => setNewAlert(prev => ({ ...prev, notificationType: value as NotificationType }))}
                    options={notificationOptions}
                    description={strings.CreateAlertNotificationDescription}
                  />
                )}
              </SharePointSection>

              <SharePointSection title={strings.CreateAlertActionLinkSectionTitle}>
                <SharePointInput
                  label={strings.CreateAlertLinkUrlLabel}
                  value={newAlert.linkUrl}
                  onChange={(value) => {
                    setNewAlert(prev => ({ ...prev, linkUrl: value }));
                    setErrors(prev => prev.linkUrl ? { ...prev, linkUrl: undefined } : prev);
                  }}
                  placeholder={strings.CreateAlertLinkUrlPlaceholder}
                  error={errors.linkUrl}
                  description={strings.CreateAlertLinkUrlDescription}
                />

                {newAlert.linkUrl && !useMultiLanguage && (
                  <SharePointInput
                    label={strings.CreateAlertLinkDescriptionLabel}
                    value={newAlert.linkDescription}
                    onChange={(value) => {
                      setNewAlert(prev => ({ ...prev, linkDescription: value }));
                      setErrors(prev => prev.linkDescription ? { ...prev, linkDescription: undefined } : prev);
                    }}
                    placeholder={strings.CreateAlertLinkDescriptionPlaceholder}
                    required={!!newAlert.linkUrl}
                    error={errors.linkDescription}
                    description={strings.CreateAlertLinkDescriptionDescription}
                  />
                )}
                
                {newAlert.linkUrl && useMultiLanguage && (
                  <div className={styles.infoMessage}>
                    <p>{strings.CreateAlertLinkDescriptionInfo}</p>
                  </div>
                )}
              </SharePointSection>

              {enableTargetSite && (
                <SharePointSection title={strings.CreateAlertTargetSitesSectionTitle}>
                  <SiteSelector
                    selectedSites={newAlert.targetSites}
                    onSitesChange={(sites) => {
                      setNewAlert(prev => ({ ...prev, targetSites: sites }));
                      if (errors.targetSites) setErrors(prev => ({ ...prev, targetSites: undefined }));
                    }}
                    siteDetector={siteDetector}
                    graphClient={graphClient}
                    showPermissionStatus={true}
                  />
                  {errors.targetSites && (
                    <div className={styles.errorMessage}>{errors.targetSites}</div>
                  )}
                </SharePointSection>
              )}

              <SharePointSection title={strings.CreateAlertSchedulingSectionTitle}>
                <SharePointInput
                  label={strings.CreateAlertStartDateLabel}
                  type="datetime-local"
                  value={DateUtils.toDateTimeLocalValue(newAlert.scheduledStart)}
                  onChange={(value) => {
                    setNewAlert(prev => ({ 
                      ...prev, 
                      scheduledStart: value ? new Date(value) : undefined 
                    }));
                    if (errors.scheduledStart) setErrors(prev => ({ ...prev, scheduledStart: undefined }));
                  }}
                  error={errors.scheduledStart}
                  description={strings.CreateAlertStartDateDescription}
                />

                <SharePointInput
                  label={strings.CreateAlertEndDateLabel}
                  type="datetime-local"
                  value={DateUtils.toDateTimeLocalValue(newAlert.scheduledEnd)}
                  onChange={(value) => {
                    setNewAlert(prev => ({ 
                      ...prev, 
                      scheduledEnd: value ? new Date(value) : undefined 
                    }));
                    if (errors.scheduledEnd) setErrors(prev => ({ ...prev, scheduledEnd: undefined }));
                  }}
                  error={errors.scheduledEnd}
                  description={strings.CreateAlertEndDateDescription}
                />
              </SharePointSection>

              <div className={styles.formActions}>
                <SharePointButton
                  variant="primary"
                  onClick={handleCreateAlert}
                  disabled={isCreatingAlert || alertTypes.length === 0}
                  icon={<Save24Regular />}
                >
                  {isCreatingAlert ? strings.CreateAlertPrimaryButtonLoading : strings.CreateAlert}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={handleSaveAsDraft}
                  disabled={isCreatingAlert || alertTypes.length === 0}
                  icon={<Drafts24Regular />}
                >
                  {strings.CreateAlertSaveDraftButtonLabel}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={resetForm}
                  disabled={isCreatingAlert}
                  icon={<Dismiss24Regular />}
                >
                  {strings.CreateAlertResetFormButtonLabel}
                </SharePointButton>

                <SharePointButton
                  variant="secondary"
                  onClick={() => setShowPreview(!showPreview)}
                  icon={<Eye24Regular />}
                >
                  {showPreview ? strings.CreateAlertHidePreview : strings.CreateAlertShowPreview}
                </SharePointButton>

                {/* Auto-save indicator */}
                <div style={{ marginLeft: 'auto', fontSize: '12px', color: '#666' }}>
                  {isAutoSaving && <span>{strings.CreateAlertAutoSaving}</span>}
                  {!isAutoSaving && lastAutoSave && (
                    <span>{Text.format(strings.CreateAlertAutoSavedAt, lastAutoSave.toLocaleTimeString())}</span>
                  )}
                </div>
              </div>

              {/* Creation Progress */}
              {creationProgress.length > 0 && (
                <div className={styles.alertsList}>
                  <h3>{strings.CreateAlertCreationResultsHeading}</h3>
                  {creationProgress.map((result, index) => (
                    <div
                      key={index}
                      className={`${styles.alertCard} ${result.error ? styles.error : styles.success}`}
                    >
                      <strong>{result.siteName}</strong>: {result.error
                        ? Text.format(strings.CreateAlertCreationResultError, result.error)
                        : strings.CreateAlertCreationResultSuccess}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Preview Column */}
            {showPreview && (
              <div className={styles.formColumn}>
                <div className={styles.alertCard}>                  
                  {/* Multi-language preview mode selector */}
                  {useMultiLanguage && newAlert.languageContent.length > 0 && (
                    <div className={styles.previewLanguageSelector}>
                      <label className={styles.previewLabel}>{strings.CreateAlertPreviewLanguageLabel}</label>
                      <div className={styles.previewLanguageButtons}>
                        {newAlert.languageContent.map((content, index) => {
                          const lang = supportedLanguages.find(l => l.code === content.language);
                          return (
                            <SharePointButton
                              key={content.language}
                              variant={index === 0 ? "primary" : "secondary"}
                              onClick={() => {
                                // Move selected language to front for preview
                                const reorderedContent = [content, ...newAlert.languageContent.filter((_, i) => i !== index)];
                                setNewAlert(prev => ({ ...prev, languageContent: reorderedContent }));
                              }}
                              className={styles.previewLanguageButton}
                            >
                              {lang?.flag || content.language} {lang?.nativeName || content.language}
                            </SharePointButton>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  <AlertPreview
                    title={useMultiLanguage && newAlert.languageContent.length > 0 
                      ? newAlert.languageContent[0]?.title || strings.CreateAlertMultiLanguagePreviewTitle
                      : newAlert.title || strings.AlertPreviewDefaultTitle}
                    description={useMultiLanguage && newAlert.languageContent.length > 0
                      ? newAlert.languageContent[0]?.description || strings.CreateAlertMultiLanguagePreviewDescription
                      : newAlert.description || strings.AlertPreviewDefaultDescription}
                    alertType={getCurrentAlertType() || { name: strings.AlertTypeInfo, iconName: "Info", backgroundColor: "#0078d4", textColor: "#ffffff", additionalStyles: "", priorityStyles: {} }}
                    priority={newAlert.priority}
                    linkUrl={newAlert.linkUrl}
                    linkDescription={useMultiLanguage && newAlert.languageContent.length > 0
                      ? newAlert.languageContent[0]?.linkDescription || strings.CreateAlertLinkPreviewFallback
                      : newAlert.linkDescription || strings.CreateAlertLinkPreviewFallback}
                    isPinned={newAlert.isPinned}
                  />

                  {/* Multi-language preview info */}
                  {useMultiLanguage && newAlert.languageContent.length > 0 && (
                    <div className={styles.multiLanguagePreviewInfo}>
                      <p><strong>{strings.CreateAlertMultiLanguagePreviewHeading}</strong></p>
                      <p>{Text.format(strings.CreateAlertMultiLanguagePreviewCurrentLanguage,
                        supportedLanguages.find(l => l.code === newAlert.languageContent[0]?.language)?.nativeName || newAlert.languageContent[0]?.language
                      )}</p>
                      <p>{Text.format(
                        strings.CreateAlertMultiLanguagePreviewAvailableLanguages,
                        newAlert.languageContent.length,
                        newAlert.languageContent.map(c => supportedLanguages.find(l => l.code === c.language)?.flag || c.language).join(' ')
                      )}</p>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default CreateAlertTab;
