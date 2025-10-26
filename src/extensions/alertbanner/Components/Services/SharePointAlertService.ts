import { MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { AlertPriority, NotificationType, IAlertType, IPersonField, ContentType, TargetLanguage } from "../Alerts/IAlerts";
import { logger } from './LoggerService';
import { AlertTransformers } from '../Utils/AlertTransformers';
import { DateUtils } from '../Utils/DateUtils';
import { LIST_NAMES, VALIDATION_LIMITS, API_CONFIG } from '../Utils/AppConstants';
import { JsonUtils } from '../Utils/JsonUtils';
import { ErrorUtils } from '../Utils/ErrorUtils';
import { AlertFilters } from '../Utils/AlertFilters';
import { RetryUtils } from '../Utils/RetryUtils';
import { StringUtils } from '../Utils/StringUtils';

export interface IRepairResult {
  success: boolean;
  message: string;
  details: {
    columnsRemoved: string[];
    columnsAdded: string[];
    columnsUpdated: string[];
    errors: string[];
    warnings: string[];
  };
}

export interface IAlertItem {
  id: string;
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  targetUsers?: IPersonField[]; // People/Groups who can see this alert. If empty, everyone sees it
  notificationType: NotificationType;
  linkUrl?: string;
  linkDescription?: string;
  targetSites: string[];
  status: 'Active' | 'Expired' | 'Scheduled' | 'Draft';
  createdDate: string;
  createdBy: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  metadata?: any;
  // New language and classification properties
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageGroup?: string;
  availableForAll?: boolean;
  // Attachments support
  attachments?: {
    fileName: string;
    serverRelativeUrl: string;
    size?: number;
  }[];
  // Store the original SharePoint list item for multi-language access
  _originalListItem?: IAlertListItem;
}

export interface IMultiLanguageContent {
  [languageCode: string]: string;
}

export interface IAlertListItem {
  Id: number;
  Title: string;
  Description: string;
  AlertType: string;
  Priority: string;
  IsPinned: boolean;
  NotificationType: string;
  LinkUrl?: string;
  LinkDescription?: string;
  TargetSites: string;
  Status: string;
  Created: string;
  Author: {
    Title: string;
  };
  ScheduledStart?: string;
  ScheduledEnd?: string;
  Metadata?: string;

  // Multi-language content fields
  Title_EN?: string;
  Title_FR?: string;
  Title_DE?: string;
  Title_ES?: string;
  Title_SV?: string;
  Title_FI?: string;
  Title_DA?: string;
  Title_NO?: string;

  Description_EN?: string;
  Description_FR?: string;
  Description_DE?: string;
  Description_ES?: string;
  Description_SV?: string;
  Description_FI?: string;
  Description_DA?: string;
  Description_NO?: string;

  LinkDescription_EN?: string;
  LinkDescription_FR?: string;
  LinkDescription_DE?: string;
  LinkDescription_ES?: string;
  LinkDescription_SV?: string;
  LinkDescription_FI?: string;
  LinkDescription_DA?: string;
  LinkDescription_NO?: string;

  // Targeting
  TargetUsers?: any[]; // SharePoint People/Groups field data

  // Language and classification properties
  ItemType?: string;
  TargetLanguage?: string;
  LanguageGroup?: string;
  AvailableForAll?: boolean;

  // Dynamic language support - for additional languages
  [key: string]: any;
}

export class SharePointAlertService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private alertsListName = LIST_NAMES.ALERTS;
  private alertTypesListName = LIST_NAMES.ALERT_TYPES;
  private listIdCache: Map<string, string> = new Map();
  private graphSiteIdentifierCache: Map<string, string> = new Map();
  private validatedListSchemas: Set<string> = new Set();

  constructor(graphClient: MSGraphClientV3, context: ApplicationCustomizerContext) {
    this.graphClient = graphClient;
    this.context = context;
  }

  private getGraphSiteIdentifierFromContext(siteId: string): string {
    const currentUrl = new URL(this.context.pageContext.web.absoluteUrl);
    const siteGuid = (siteId || this.context.pageContext.site.id.toString()).replace(/[{}]/g, '');
    const webGuid = this.context.pageContext.web.id.toString().replace(/[{}]/g, '');
    return `${currentUrl.hostname},${siteGuid},${webGuid}`;
  }

  private isCurrentSite(siteId: string): boolean {
    if (!siteId) {
      return true;
    }

    const normalized = siteId.replace(/[{}]/g, '').toLowerCase();
    const currentSiteId = this.context.pageContext.site.id.toString().replace(/[{}]/g, '').toLowerCase();
    return normalized === currentSiteId;
  }

  private async ensureGraphSiteIdentifier(siteId: string): Promise<string> {
    if (!siteId) {
      return this.getGraphSiteIdentifierFromContext(siteId);
    }

    if (siteId.includes(',')) {
      return siteId;
    }

    if (siteId.startsWith('https://') || siteId.startsWith('http://')) {
      try {
        const siteUrl = new URL(siteId);
        const rawPath = siteUrl.pathname || '';
        const normalizedPath = rawPath.replace(/\/+/g, '/').replace(/\/$/, '');
        return normalizedPath ? `${siteUrl.hostname}:${normalizedPath || '/'}` : siteUrl.hostname;
      } catch (error) {
        logger.warn('SharePointAlertService', 'Invalid site URL provided, falling back to context identifier', { siteId, error });
        return this.getGraphSiteIdentifierFromContext(siteId);
      }
    }

    if (siteId.includes(':')) {
      return siteId;
    }

    const normalized = siteId.replace(/[{}]/g, '').toLowerCase();
    if (this.graphSiteIdentifierCache.has(normalized)) {
      return this.graphSiteIdentifierCache.get(normalized)!;
    }

    if (this.isCurrentSite(siteId)) {
      const identifier = this.getGraphSiteIdentifierFromContext(siteId);
      this.graphSiteIdentifierCache.set(normalized, identifier);
      return identifier;
    }

    try {
      const siteResponse = await this.graphClient
        .api(`/sites/${normalized}`)
        .select('id')
        .get();

      if (siteResponse?.id) {
        this.graphSiteIdentifierCache.set(normalized, siteResponse.id);
        return siteResponse.id;
      }
    } catch (error) {
      logger.warn('SharePointAlertService', 'Unable to resolve graph site identifier, falling back to context derived value', { siteId, error });
    }

    const fallback = this.getGraphSiteIdentifierFromContext(siteId);
    this.graphSiteIdentifierCache.set(normalized, fallback);
    return fallback;
  }

  private getListCacheKey(siteIdentifier: string, listTitle: string): string {
    return `${siteIdentifier}|${listTitle.toLowerCase()}`;
  }


  private async resolveListId(siteId: string, listTitle: string): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const cacheKey = this.getListCacheKey(graphSiteIdentifier, listTitle);
    const cachedId = this.listIdCache.get(cacheKey);
    if (cachedId) {
      return cachedId;
    }

    try {
      let requestPath = `/sites/${graphSiteIdentifier}/lists`;
      const targetName = listTitle.toLowerCase();

      while (requestPath) {
        const response = await this.graphClient
          .api(requestPath)
          .select('id,displayName,name')
          .top(200)
          .get();

        const match = response.value?.find((list: any) => {
          const displayName = (list.displayName || '').toLowerCase();
          const name = (list.name || '').toLowerCase();
          return displayName === targetName || name === targetName;
        });

        if (match?.id) {
          this.listIdCache.set(cacheKey, match.id);
          return match.id;
        }

        if (response['@odata.nextLink']) {
          requestPath = response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '');
        } else {
          requestPath = '';
        }
      }
    } catch (error) {
      logger.warn('SharePointAlertService', `Unable to enumerate lists while resolving ${listTitle}`, { siteId, error });
      throw error;
    }

    const notFoundError = new Error(`LIST_NOT_FOUND:${listTitle}`);
    notFoundError.name = 'LIST_NOT_FOUND';
    throw notFoundError;
  }

  private async registerListId(siteId: string, listTitle: string, listId?: string): Promise<void> {
    if (!listId) {
      return;
    }

    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    this.listIdCache.set(this.getListCacheKey(graphSiteIdentifier, listTitle), listId);
  }

  private async getAlertsListApi(siteId: string): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const listId = await this.resolveListId(siteId, this.alertsListName);
    return `/sites/${graphSiteIdentifier}/lists/${listId}`;
  }

  private async getAlertTypesListApi(siteId: string): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const listId = await this.resolveListId(siteId, this.alertTypesListName);
    return `/sites/${graphSiteIdentifier}/lists/${listId}`;
  }


  /**
   * Execute SharePoint API call with retry logic for transient failures
   */
  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    maxRetries: number = 3,
    baseDelay: number = 1000
  ): Promise<T> {
    return RetryUtils.executeWithRetry(operation, {
      maxRetries,
      baseDelay,
      useExponentialBackoff: true,
      useJitter: true,
      shouldRetry: (error) => ErrorUtils.isRetryableError(error)
    });
  }


  /**
   * Check if the current site is the SharePoint home site
   */
  private async isHomeSite(siteId: string): Promise<boolean> {
    try {
      // Get the SharePoint home site ID
      const homeSiteResponse = await this.graphClient
        .api("/sites/root")
        .select("id")
        .get();
      const homeSiteId: string = homeSiteResponse.id;

      return siteId === homeSiteId;
    } catch (error) {
      logger.warn('SharePointAlertService', 'Unable to determine if site is home site, assuming it is not', error);
      return false;
    }
  }

  /**
   * Initialize SharePoint lists if they don't exist
   */
  /**
   * Check which sites need list creation
   */
  public async checkListsNeeded(): Promise<{ site: string; needsAlerts: boolean; needsTypes: boolean; isHomeSite: boolean }[]> {
    const results = [];
    const currentSiteId = this.context.pageContext.site.id.toString();
    
    // Check if current site is home site
    const isHomeSite = await this.isHomeSite(currentSiteId);
    
    // Check current site
    let needsAlerts = false;
    let needsTypes = false;
    
    try {
      await this.resolveListId(currentSiteId, this.alertsListName);
    } catch (error: any) {
      if (ErrorUtils.isListNotFoundError(error)) {
        needsAlerts = true;
      } else if (!ErrorUtils.isAccessDeniedError(error)) {
        throw error;
      }
    }

    // Only check for AlertBannerTypes if this is the home site
    if (isHomeSite) {
      try {
        await this.resolveListId(currentSiteId, this.alertTypesListName);
      } catch (error: any) {
        if (ErrorUtils.isListNotFoundError(error)) {
          needsTypes = true;
        } else if (!ErrorUtils.isAccessDeniedError(error)) {
          throw error;
        }
      }
    }
    
    results.push({
      site: currentSiteId,
      needsAlerts,
      needsTypes,
      isHomeSite
    });
    
    return results;
  }

  public async initializeLists(): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const isHomeSite = await this.isHomeSite(siteId);

      try {
        await this.ensureAlertsList(siteId);
      } catch (alertsError) {
        if (alertsError.message?.includes('PERMISSION_DENIED')) {
          logger.warn('SharePointAlertService', 'Cannot create alerts list due to insufficient permissions');
        } else {
          throw alertsError;
        }
      }

      if (isHomeSite) {
        try {
          await this.ensureAlertTypesList(siteId);
        } catch (typesError) {
          if (typesError.message?.includes('PERMISSION_DENIED')) {
            logger.warn('SharePointAlertService', 'Cannot create types list on home site due to insufficient permissions');
          } else {
            throw typesError;
          }
        }
      }
    } catch (error) {
      // Enhanced error handling for common permission issues
      if (error.message?.includes('PERMISSION_DENIED')) {
        logger.warn('SharePointAlertService', 'SharePoint list creation failed due to insufficient permissions.');
        throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        logger.warn('SharePointAlertService', 'SharePoint lists not found and cannot be created.');
        throw new Error('LISTS_NOT_FOUND: SharePoint lists do not exist and cannot be created.');
      } else {
        logger.error('SharePointAlertService', 'Failed to initialize SharePoint lists', error);
        throw new Error(`INITIALIZATION_FAILED: ${error.message || 'Unknown error during SharePoint initialization'}`);
      }
    }
  }

  /**
   * Create alerts list if it doesn't exist
   */
  private async ensureAlertsList(siteId: string): Promise<boolean> {
    try {
      await this.resolveListId(siteId, this.alertsListName);
      return false;
    } catch (error: any) {
      if (ErrorUtils.isAccessDeniedError(error)) {
        logger.warn('SharePointAlertService', 'Cannot access or create alerts list due to insufficient permissions');
        throw new Error('PERMISSION_DENIED: User lacks permissions to access or create SharePoint lists.');
      }

      if (!ErrorUtils.isListNotFoundError(error)) {
        throw error;
      }
    }

    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);

    try {
      await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .select('id')
        .top(1)
        .get();
    } catch (permissionError) {
      const errorMessage = permissionError.message || '';
      const statusCode = permissionError.code || '';

      logger.error('SharePointAlertService', 'Permission check failed', permissionError, {
        message: errorMessage,
        code: statusCode,
        siteId
      });

      if (errorMessage.includes('Access denied') || statusCode === '403' || errorMessage.includes('403')) {
        throw new Error('PERMISSION_DENIED: User lacks Sites.ReadWrite.All permissions to create SharePoint lists. Please contact your SharePoint administrator to grant the required permissions.');
      } else if (statusCode === '401') {
        throw new Error('AUTHENTICATION_FAILED: User authentication failed. Please re-authenticate.');
      } else {
        throw new Error(`PERMISSION_CHECK_FAILED: Unable to verify permissions - ${errorMessage}`);
      }
    }

    const listDefinition = {
      displayName: this.alertsListName,
      list: {
        template: 'genericList',
        contentTypesEnabled: false
      }
    };

    try {
      const createdList = await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .post(listDefinition);
      await this.registerListId(siteId, this.alertsListName, createdList?.id);

      await this.enableListAttachments(siteId, createdList?.id);
      await this.addAlertsListColumns(siteId);
      await this.seedDefaultAlertTypes(siteId);
      await this.createTemplateAlerts(siteId);

      return true;
    } catch (createError) {
      if (ErrorUtils.isAccessDeniedError(createError)) {
        logger.warn('SharePointAlertService', 'User lacks permissions to create SharePoint lists');
        throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
      }
      if (createError.message?.includes('CRITICAL_COLUMNS_FAILED')) {
        logger.error('SharePointAlertService', 'List created but critical columns failed', createError);
        throw new Error(`LIST_INCOMPLETE: ${createError.message}`);
      }
      throw createError;
    }
  }

  /**
   * Enable attachments on the Alerts list
   */
  private async enableListAttachments(siteId: string, listId: string): Promise<void> {
    try {
      const siteUrl = await this.getSiteUrlFromIdentifier(siteId);
      const updateUrl = `${siteUrl}/_api/web/lists(guid'${listId}')`;

      await this.context.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'EnableAttachments': true
          })
        }
      );

      logger.info('SharePointAlertService', 'Attachments enabled on Alerts list');
    } catch (error) {
      logger.warn('SharePointAlertService', 'Failed to enable attachments (non-critical)', error);
      // Don't throw - attachments are optional functionality
    }
  }

  /**
   * Add custom columns to the Alerts list after creation
   */
  private async addAlertsListColumns(siteId: string): Promise<void> {
    let alertTypesListId = '';
    try {
      alertTypesListId = await this.resolveListId(siteId, this.alertTypesListName);
    } catch (error) {
      logger.warn('SharePointAlertService', 'Could not get AlertBannerTypes list ID for lookup field', error);
    }

    const alertsListId = await this.resolveListId(siteId, this.alertsListName);

    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);

    const columns = [
      alertTypesListId ? {
        name: 'AlertType',
        lookup: {
          listId: alertTypesListId,
          columnName: 'Title',
          allowMultipleValues: false,
          allowUnlimitedLength: false
        }
      } : {
        name: 'AlertType',
        text: {
          maxLength: 255
        }
      },
      {
        name: 'Description',
        text: {
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'Priority',
        choice: {
          allowTextEntry: false,
          choices: ['low', 'medium', 'high', 'critical'],
          displayAs: 'dropdown'
        }
      },
      {
        name: 'IsPinned',
        boolean: {}
      },
      {
        name: 'NotificationType',
        choice: {
          allowTextEntry: false,
          choices: ['none', 'browser', 'email', 'both'],
          displayAs: 'dropdown'
        }
      },
      {
        name: 'LinkUrl',
        text: {}
      },
      {
        name: 'LinkDescription',
        text: {
          maxLength: 255
        }
      },
      {
        name: 'TargetSites',
        text: {
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'Status',
        choice: {
          allowTextEntry: false,
          choices: ['Active', 'Expired', 'Scheduled'],
          displayAs: 'dropdown'
        }
      },
      {
        name: 'ScheduledStart',
        dateTime: {
          displayAs: 'default',
          format: 'dateTime'
        }
      },
      {
        name: 'ScheduledEnd',
        dateTime: {
          displayAs: 'default',
          format: 'dateTime'
        }
      },
      {
        name: 'Metadata',
        text: {
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'ItemType',
        choice: {
          allowTextEntry: false,
          choices: ['alert', 'template'],
          displayAs: 'dropdown'
        }
      },
      {
        name: 'TargetLanguage',
        choice: {
          allowTextEntry: false,
          choices: ['all', 'en-us', 'fr-fr', 'de-de', 'es-es', 'sv-se', 'fi-fi', 'da-dk', 'nb-no'],
          displayAs: 'dropdown'
        }
      },
      {
        name: 'LanguageGroup',
        text: {
          maxLength: 255
        }
      },
      {
        name: 'AvailableForAll',
        boolean: {}
      },
      {
        name: 'TargetUsers',
        personOrGroup: {
          allowMultipleSelection: true,
          chooseFromType: 'peopleAndGroups'
        }
      }
    ];

    for (const column of columns) {
      try {
        await this.graphClient
          .api(`/sites/${graphSiteIdentifier}/lists/${alertsListId}/columns`)
          .post(column);
      } catch (error) {
        logger.warn('SharePointAlertService', `Failed to create Alerts column ${column.name}`, error);
        if (column.name === 'AlertType') {
          logger.error('SharePointAlertService', 'CRITICAL_COLUMNS_FAILED: AlertType column creation failed', error);
          throw new Error('CRITICAL_COLUMNS_FAILED: Failed to create AlertType lookup column');
        }
      }
    }
  }

  private async getSiteUrlFromIdentifier(siteId: string): Promise<string> {
    if (!siteId) {
      return this.context.pageContext.web.absoluteUrl;
    }

    if (siteId.startsWith('https://')) {
      return siteId;
    }

    if (siteId.includes(':') && !siteId.includes(',')) {
      const [hostname, path = '/'] = siteId.split(':');
      return `https://${hostname}${path === '/' ? '' : path}`;
    }

    if (siteId.includes(',')) {
      try {
        const siteDetails = await this.graphClient
          .api(`/sites/${siteId}`)
          .select('webUrl')
          .get();

        if (siteDetails?.webUrl) {
          return siteDetails.webUrl;
        }
      } catch (error) {
        logger.warn('SharePointAlertService', 'Unable to resolve site URL from identifier', { siteId, error });
      }
    }

    return this.context.pageContext.web.absoluteUrl;
  }

  /**
   * Create template alert items when list is first created
   */
  private async createTemplateAlerts(siteId: string): Promise<void> {
    // Import template data from JSON file
    const defaultTemplates = require('../Data/defaultTemplates.json');
    
    // Add dynamic dates to templates and map ContentType to ItemType
    const templateAlerts = defaultTemplates.map((template: any) => ({
      ...template,
      fields: {
        ...template.fields,
        ScheduledStart: new Date().toISOString(),
        // Set different end dates based on alert type for variety
        ScheduledEnd: this.getTemplateEndDate(template.fields.AlertType),
        // Map ContentType to ItemType for SharePoint
        ItemType: template.fields.ContentType,
        // Remove ContentType as it's not a SharePoint column
        ContentType: undefined
      }
    }));

    const alertsListApi = await this.getAlertsListApi(siteId);

    for (const template of templateAlerts) {
      try {
        await this.graphClient
          .api(`${alertsListApi}/items`)
          .post(template);
        logger.debug('SharePointAlertService', `Created template: ${template.fields.Title}`);
      } catch (error) {
        logger.warn('SharePointAlertService', `Failed to create template: ${template.fields.Title}`, error);
        // Don't throw error for template creation failures - they're nice-to-have
      }
    }
  }

  /**
   * Get appropriate end date for template based on alert type using DateUtils
   */
  private getTemplateEndDate(alertType: string): string {
    const now = new Date();
    switch (alertType.toLowerCase()) {
      case 'maintenance':
        return DateUtils.addDurationISO(now, 1, 'days');
      case 'warning':
        return DateUtils.addDurationISO(now, 3, 'days');
      case 'interruption':
        return DateUtils.addDurationISO(now, 12, 'hours');
      case 'info':
        return DateUtils.addDurationISO(now, 1, 'weeks');
      default:
        return DateUtils.addDurationISO(now, 1, 'months');
    }
  }

  /**
   * Get template alerts for the AlertTemplates component
   */
  public async getTemplateAlerts(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.getAlertsListApi(siteId);
      const response = await this.executeWithRetry(() =>
        this.graphClient
          .api(`${alertsListApi}/items`)
          .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
          .expand("fields($select=Title,AlertType,Description,Priority,IsPinned,NotificationType,LinkUrl,LinkDescription,TargetSites,Status,ItemType,TargetLanguage,LanguageGroup,ScheduledStart,ScheduledEnd,TargetUsers)")
          .filter("fields/ItemType eq 'template'")
          .orderby("fields/Created desc")
          .get()
      );

      return response.value.map((item: any) => this.mapSharePointItemToAlert(item));
    } catch (error) {
      logger.warn('SharePointAlertService', 'Could not fetch template alerts after retries', error);
      return [];
    }
  }

  /**
   * Get draft alerts for the current user
   */
  public async getDraftAlerts(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.getAlertsListApi(siteId);
      const currentUser = this.context.pageContext.user.loginName;

      const response = await this.executeWithRetry(() =>
        this.graphClient
          .api(`${alertsListApi}/items`)
          .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
          .expand("fields($select=Title,AlertType,Description,Priority,IsPinned,NotificationType,LinkUrl,LinkDescription,TargetSites,Status,ItemType,TargetLanguage,LanguageGroup,ScheduledStart,ScheduledEnd,TargetUsers,Author,Modified)")
          .filter(`fields/ItemType eq 'draft' and fields/Author/Email eq '${currentUser}'`)
          .orderby("fields/Modified desc")
          .get()
      );

      return response.value.map((item: any) => this.mapSharePointItemToAlert(item));
    } catch (error) {
      logger.warn('SharePointAlertService', 'Could not fetch draft alerts after retries', error);
      return [];
    }
  }

  /**
   * Save alert as draft
   */
  public async saveDraft(draft: Partial<IAlertItem>): Promise<IAlertItem> {
    const alertData: any = {
      ...draft,
      contentType: ContentType.Draft,
      status: 'Draft'
    };

    if (draft.id && parseInt(draft.id) > 0) {
      // Update existing draft
      return this.updateAlert(draft.id, alertData);
    } else {
      // Create new draft
      return this.createAlert(alertData);
    }
  }

  /**
   * Delete draft alert
   */
  public async deleteDraft(draftId: string): Promise<void> {
    return this.deleteAlert(draftId);
  }

  /**
   * Create alert types list if it doesn't exist
   */
  private async ensureAlertTypesList(siteId: string): Promise<boolean> {
    try {
      await this.resolveListId(siteId, this.alertTypesListName);
      return false;
    } catch (error: any) {
      if (ErrorUtils.isAccessDeniedError(error)) {
        logger.warn('SharePointAlertService', 'Cannot access or create alert types list due to insufficient permissions');
        throw new Error('PERMISSION_DENIED: User lacks permissions to access or create SharePoint lists.');
      }

      if (!ErrorUtils.isListNotFoundError(error)) {
        throw error;
      }
    }

    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);

    try {
      await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .select('id')
        .top(1)
        .get();
    } catch (permissionError) {
      if (ErrorUtils.isAccessDeniedError(permissionError)) {
        logger.warn('SharePointAlertService', 'User lacks permissions to create SharePoint lists');
        throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
      }
    }

    logger.info('SharePointAlertService', 'Creating alert types list');

    const listDefinition = {
      displayName: this.alertTypesListName,
      list: {
        template: 'genericList'
      }
    };

    try {
      const createdList = await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .post(listDefinition);
      await this.registerListId(siteId, this.alertTypesListName, createdList?.id);

      await this.addAlertTypesListColumns(siteId);
      await this.seedDefaultAlertTypes(siteId);

      return true;
    } catch (createError) {
      if (ErrorUtils.isAccessDeniedError(createError)) {
        logger.warn('SharePointAlertService', 'User lacks permissions to create SharePoint lists');
        throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
      }
      throw createError;
    }
  }

  /**
   * Add custom columns to the AlertTypes list after creation
   */
  private async addAlertTypesListColumns(siteId: string): Promise<void> {
    const columns = [
      {
        name: 'IconName',
        text: {
          maxLength: 100,
          allowMultipleLines: false
        }
      },
      {
        name: 'BackgroundColor',
        text: { 
          maxLength: 50,
          allowMultipleLines: false
        }
      },
      {
        name: 'TextColor',
        text: { 
          maxLength: 50,
          allowMultipleLines: false
        }
      },
      {
        name: 'AdditionalStyles',
        text: { 
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'PriorityStyles',
        text: { 
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'SortOrder',
        number: { 
          decimalPlaces: 'none'
        },
        indexed: true
      }
    ];

    const alertTypesListApi = await this.getAlertTypesListApi(siteId);

    for (const column of columns) {
      try {
        await this.graphClient
          .api(`${alertTypesListApi}/columns`)
          .post(column);
      } catch (error) {
        logger.warn('SharePointAlertService', `Failed to create AlertTypes column ${column.name}`, error);
        // Continue creating other columns even if one fails
      }
    }
  }

  private async seedDefaultAlertTypes(siteId: string): Promise<void> {
    try {
      const alertTypesListApi = await this.getAlertTypesListApi(siteId);
      const existing = await this.graphClient
        .api(`${alertTypesListApi}/items`)
        .top(1)
        .get();

      if (existing.value && existing.value.length > 0) {
        return;
      }

      const defaults = this.getDefaultAlertTypes();
      let sortOrder = 0;
      for (const alertType of defaults) {
        const payload = {
          fields: {
            Title: alertType.name,
            IconName: alertType.iconName,
            BackgroundColor: alertType.backgroundColor,
            TextColor: alertType.textColor,
            AdditionalStyles: alertType.additionalStyles || '',
            PriorityStyles: JsonUtils.safeStringify(alertType.priorityStyles || {}) || '{}',
            SortOrder: sortOrder++
          }
        };

        try {
          await this.graphClient
            .api(`${alertTypesListApi}/items`)
            .post(payload);
        } catch (error) {
          logger.warn('SharePointAlertService', 'Failed to seed default alert type', { name: alertType.name, error });
        }
      }

    } catch (error) {
      logger.warn('SharePointAlertService', 'Unable to seed default alert types', { siteId, error });
    }
  }

  /**
   * Get all alerts from SharePoint
   */
  public async getAlerts(siteIds?: string[]): Promise<IAlertItem[]> {
    try {
      let sitesToQuery = siteIds;
      
      // If no specific sites provided, use hierarchical sites from SiteContextService
      if (!sitesToQuery) {
        try {
          // Import dynamically to avoid circular dependency
          const { SiteContextService } = await import('./SiteContextService');
          const siteContextService = SiteContextService.getInstance(this.context, this.graphClient);
          await siteContextService.initialize();
          sitesToQuery = siteContextService.getAlertSourceSites();
        } catch (error) {
          logger.warn('SharePointAlertService', 'Failed to get hierarchical sites, falling back to current site', error);
          sitesToQuery = [this.context.pageContext.site.id.toString()];
        }
      }
      const dedupMap = new Map<string, string>();
      sitesToQuery.forEach(siteId => {
        const normalized = siteId.includes(',')
          ? siteId.split(',')[1] || siteId
          : siteId.replace(/[{}]/g, '').toLowerCase();
        if (!dedupMap.has(normalized)) {
          dedupMap.set(normalized, siteId);
        }
      });

      const uniqueSiteIds = Array.from(dedupMap.values());
      const allAlerts: IAlertItem[] = [];
      const batchSize = 3;

      for (let i = 0; i < uniqueSiteIds.length; i += batchSize) {
        const batch = uniqueSiteIds.slice(i, i + batchSize);
        const batchResults = await Promise.allSettled(batch.map(siteId => this.fetchAlertsForSite(siteId)));

        batchResults.forEach((result, index) => {
          if (result.status === 'fulfilled') {
            allAlerts.push(...result.value);
          } else {
            logger.warn('SharePointAlertService', `Failed to get alerts from site ${batch[index]}`, result.reason);
          }
        });
      }

      // Remove duplicates and sort by creation date
      const uniqueAlerts = AlertFilters.removeDuplicates(allAlerts);
      return uniqueAlerts.sort((a, b) => new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime());
    } catch (error) {
      // Enhanced error handling for permission and access issues
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        logger.warn('SharePointAlertService', 'Access denied when trying to get alerts from SharePoint');
        throw new Error('PERMISSION_DENIED: Cannot access SharePoint alerts due to insufficient permissions.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        logger.warn('SharePointAlertService', 'SharePoint alerts list not found');
        throw new Error('LISTS_NOT_FOUND: SharePoint alerts list does not exist.');
      } else {
        logger.error('SharePointAlertService', 'Failed to get alerts', error);
        throw new Error(`GET_ALERTS_FAILED: ${error.message || 'Unknown error when retrieving alerts'}`);
      }
    }
  }

  private async fetchAlertsForSite(siteId: string): Promise<IAlertItem[]> {
    try {
      const alertsListApi = await this.getAlertsListApi(siteId);
      const response = await this.graphClient
        .api(`${alertsListApi}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand("fields($select=Title,AlertType,Description,Priority,IsPinned,NotificationType,LinkUrl,LinkDescription,TargetSites,Status,ItemType,TargetLanguage,LanguageGroup,ScheduledStart,ScheduledEnd,TargetUsers,Created,Author,Attachments,AttachmentFiles)")
        .orderby("fields/Created desc")
        .get();

      return response.value
        .filter((item: any) => {
          const title = item.fields?.Title || '';
          const itemType = (item.fields?.ItemType || '').toLowerCase();
          const status = (item.fields?.Status || '').toLowerCase();

          return itemType !== 'draft' &&
            itemType !== 'template' &&
            status !== 'draft' &&
            !title.startsWith('[Auto-saved]') &&
            !title.startsWith('[auto-saved]');
        })
        .map((item: any) => this.mapSharePointItemToAlert(item, siteId));
    } catch (error) {
      logger.warn('SharePointAlertService', `Failed to get alerts from site ${siteId}`, error);
      return [];
    }
  }

  /**
   * Create a new alert
   */
  public async createAlert(alert: Omit<IAlertItem, 'id' | 'createdDate' | 'createdBy' | 'status'> & Partial<Pick<IAlertItem, 'status'>>): Promise<IAlertItem> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      // Validate required fields
      if (!alert.title?.trim()) {
        throw new Error('Alert title is required');
      }
      if (!alert.description?.trim()) {
        throw new Error('Alert description is required');
      }
      if (!alert.AlertType?.trim()) {
        throw new Error('Alert type is required');
      }
      if (!alert.targetSites || alert.targetSites.length === 0) {
        throw new Error('At least one target site is required');
      }

      const alertsListApi = await this.getAlertsListApi(siteId);

      const schemaCacheKey = `${siteId}:${alertsListApi}`;
      if (!this.validatedListSchemas.has(schemaCacheKey)) {
        try {
          const listInfo = await this.graphClient
            .api(alertsListApi)
            .expand('columns')
            .get();
          
          const columnNames = listInfo.columns.map((col: any) => col.name);
          const requiredColumns = ['Title', 'Description', 'AlertType', 'Priority', 'IsPinned'];
          const missingColumns = requiredColumns.filter(col => !columnNames.includes(col));
          if (missingColumns.length > 0) {
            throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
          }

          this.validatedListSchemas.add(schemaCacheKey);
        } catch (listError: any) {
          logger.error('SharePointAlertService', 'Failed to validate list structure', listError);
          if (listError.message?.includes('Missing required columns')) {
            throw listError;
          }
          // Continue if we can't check the list structure
        }
      }

      // Build the list item carefully with proper data types
      const fields: any = {
        Title: alert.title.trim(),
        Description: alert.description.trim(),
        AlertType: alert.AlertType.trim(), // This should be the lookup value (just the text name)
        Priority: alert.priority,
        IsPinned: Boolean(alert.isPinned),
        NotificationType: alert.notificationType
      };

      // Add optional fields only if they have values
      if (alert.linkUrl?.trim()) {
        fields.LinkUrl = alert.linkUrl.trim();
      }
      if (alert.linkDescription?.trim()) {
        fields.LinkDescription = alert.linkDescription.trim();
      }
      if (alert.targetSites && alert.targetSites.length > 0) {
        const targetSitesStr = JsonUtils.safeStringify(alert.targetSites);
        if (targetSitesStr) {
          fields.TargetSites = targetSitesStr;
        } else {
          logger.error('SharePointAlertService', 'Failed to serialize targetSites', { alertId: alert.title });
        }
      }

      // Set status: use provided status or auto-determine from scheduling
      fields.Status = alert.status || (alert.scheduledStart && new Date(alert.scheduledStart) > new Date() ? 'Scheduled' : 'Active');

      if (alert.scheduledStart) {
        fields.ScheduledStart = new Date(alert.scheduledStart).toISOString();
      }
      if (alert.scheduledEnd) {
        fields.ScheduledEnd = new Date(alert.scheduledEnd).toISOString();
      }
      if (alert.metadata) {
        const metadataStr = JsonUtils.safeStringify(alert.metadata);
        if (metadataStr) {
          fields.Metadata = metadataStr;
        } else {
          logger.warn('SharePointAlertService', 'Failed to serialize metadata', { alertId: alert.title });
        }
      }

      // Add targeting
      if (alert.targetUsers && alert.targetUsers.length > 0) {
        fields.TargetUsers = alert.targetUsers;
      }

      // Add language and classification properties
      fields.ItemType = alert.contentType;
      fields.TargetLanguage = alert.targetLanguage;
      
      if (alert.languageGroup) {
        fields.LanguageGroup = alert.languageGroup;
      }
      fields.AvailableForAll = Boolean(alert.availableForAll);

      const listItem = { fields };

      logger.debug('SharePointAlertService', 'Creating alert', {
        alert,
        listItem: { ...listItem, fields: { ...listItem.fields, Description: StringUtils.truncate(listItem.fields.Description, 100) } }
      });

      let response;
      try {
        response = await this.graphClient
          .api(`${alertsListApi}/items`)
          .post(listItem);
        
        logger.debug('SharePointAlertService', 'Alert created successfully', { alertId: response.id });
      } catch (graphError: any) {
        // Parse the error object properly  
        const errorDetails = {
          message: graphError.message || 'Unknown error',
          code: graphError.code,
          statusCode: graphError.statusCode,
          body: graphError.body,
          stack: graphError.stack,
          name: graphError.name,
          fullError: JSON.stringify(graphError, Object.getOwnPropertyNames(graphError)),
          requestData: listItem
        };
        
        logger.error('SharePointAlertService', 'MS Graph API error when creating alert', errorDetails);
        
        // Try with minimal fields if the full request fails
        logger.warn('SharePointAlertService', 'Full request failed, trying with minimal fields');
        
        try {
          const minimalItem = {
            fields: {
              Title: alert.title.trim(),
              Description: alert.description.trim(),
              AlertType: alert.AlertType.trim(),
              Priority: alert.priority,
              IsPinned: Boolean(alert.isPinned),
              NotificationType: alert.notificationType,
              Status: 'Active'
            }
          };
          
          logger.debug('SharePointAlertService', 'Trying minimal request', minimalItem);
          
          response = await this.graphClient
            .api(`${alertsListApi}/items`)
            .post(minimalItem);
            
          logger.info('SharePointAlertService', 'Alert created with minimal fields', { alertId: response.id });
        } catch (minimalError: any) {
          logger.error('SharePointAlertService', 'Even minimal request failed', {
            error: minimalError.message,
            fullError: JSON.stringify(minimalError, Object.getOwnPropertyNames(minimalError))
          });
          
          // Provide more specific error message based on the error
          if (graphError.message?.includes('column') || graphError.message?.includes('field')) {
            throw new Error(`Field validation error: ${graphError.message}`);
          } else if (graphError.message?.includes('lookup')) {
            throw new Error(`Lookup field error: ${graphError.message}`);
          } else if (graphError.message?.includes('required')) {
            throw new Error(`Required field missing: ${graphError.message}`);
          }
          throw minimalError;
        }
      }

      // Get the created item with expanded fields
      try {
        const createdItem = await this.graphClient
          .api(`${alertsListApi}/items/${response.id}`)
          .expand('fields')
          .get();

        return this.mapSharePointItemToAlert(createdItem, siteId);
      } catch (retrieveError: any) {
        logger.warn('SharePointAlertService', 'Alert created but failed to retrieve details', { 
          alertId: response.id, 
          error: retrieveError.message 
        });
        // Return basic alert info if we can't retrieve the full details
        throw new Error('Alert created but could not retrieve details');
      }
    } catch (error) {
      logger.error('SharePointAlertService', 'Failed to create alert', error);
      throw error;
    }
  }

  /**
   * Extract site ID and item ID from composite alert ID
   */
  public parseAlertId(alertId: string): { siteId: string; itemId: string } {
    const lastHyphenIndex = alertId.lastIndexOf('-');
    if (lastHyphenIndex > 0 && lastHyphenIndex < alertId.length - 1) {
      const siteId = alertId.substring(0, lastHyphenIndex);
      const itemId = alertId.substring(lastHyphenIndex + 1);
      // Check if itemId is numeric (valid SharePoint item ID)
      if (/^\d+$/.test(itemId)) {
        return { siteId, itemId };
      }
    }
    // For backward compatibility, assume current site if no composite ID
    return { siteId: this.context.pageContext.site.id.toString(), itemId: alertId };
  }

  public getAlertSiteId(alertId: string): string {
    return this.parseAlertId(alertId).siteId;
  }

  /**
   * Update an existing alert
   */
  public async updateAlert(alertId: string, updates: Partial<IAlertItem>): Promise<IAlertItem> {
    try {
      const { siteId, itemId } = this.parseAlertId(alertId);
      const alertsListApi = await this.getAlertsListApi(siteId);

      const listItem = {
        fields: {
          ...(updates.title && { Title: updates.title }),
          ...(updates.description && { Description: updates.description }),
          ...(updates.AlertType && { AlertType: updates.AlertType }),
          ...(updates.priority && { Priority: updates.priority }),
          ...(updates.isPinned !== undefined && { IsPinned: updates.isPinned }),
          ...(updates.notificationType && { NotificationType: updates.notificationType }),
          ...(updates.linkUrl !== undefined && { LinkUrl: updates.linkUrl }),
          ...(updates.linkDescription !== undefined && { LinkDescription: updates.linkDescription }),
          ...(updates.targetSites && { TargetSites: JsonUtils.safeStringify(updates.targetSites) || '[]' }),
          ...(updates.scheduledStart !== undefined && { ScheduledStart: updates.scheduledStart }),
          ...(updates.scheduledEnd !== undefined && { ScheduledEnd: updates.scheduledEnd }),
          ...(updates.targetUsers !== undefined && { TargetUsers: updates.targetUsers || [] }),
          ...(updates.metadata && { Metadata: JsonUtils.safeStringify(updates.metadata) || '{}' })
        }
      };

      await this.graphClient
        .api(`${alertsListApi}/items/${itemId}/fields`)
        .patch(listItem.fields);

      // Get the updated item
      const updatedItem = await this.graphClient
        .api(`${alertsListApi}/items/${itemId}`)
        .expand('fields')
        .get();

      return this.mapSharePointItemToAlert(updatedItem, siteId);
    } catch (error) {
      logger.error('SharePointAlertService', 'Failed to update alert', error);
      throw error;
    }
  }

  /**
   * Delete an alert
   */
  public async deleteAlert(alertId: string): Promise<void> {
    try {
      const { siteId, itemId } = this.parseAlertId(alertId);
      const alertsListApi = await this.getAlertsListApi(siteId);

      await this.graphClient
        .api(`${alertsListApi}/items/${itemId}`)
        .delete();
    } catch (error) {
      logger.error('SharePointAlertService', 'Failed to delete alert', error);
      throw error;
    }
  }

  /**
   * Delete multiple alerts
   */
  public async deleteAlerts(alertIds: string[]): Promise<void> {
    const deletePromises = alertIds.map(id => this.deleteAlert(id));
    await Promise.allSettled(deletePromises);
  }

  /**
   * Get alert types from SharePoint
   */
  public async getAlertTypes(siteIdOverride?: string): Promise<IAlertType[]> {
    try {
      const siteId = siteIdOverride && siteIdOverride.trim().length > 0
        ? siteIdOverride
        : this.context.pageContext.site.id.toString();

      // Try to ensure the alert types list exists
      try {
        await this.ensureAlertTypesList(siteId);
      } catch (ensureError) {
        logger.warn('SharePointAlertService', 'Could not ensure alert types list exists', ensureError);
      }

      const alertTypesListApi = await this.getAlertTypesListApi(siteId);

      const response = await this.graphClient
        .api(`${alertTypesListApi}/items`)
        .expand('fields')
        .orderby('fields/SortOrder')
        .get();

      if (!response.value || response.value.length === 0) {
        logger.warn('SharePointAlertService', 'Alert types list is empty, seeding defaults');
        await this.seedDefaultAlertTypes(siteId);
        return this.getDefaultAlertTypes();
      }

      return response.value.map((item: any) => this.mapSharePointItemToAlertType(item));
    } catch (error) {
      logger.warn('SharePointAlertService', 'Failed to get alert types from SharePoint, using defaults', error);
      return this.getDefaultAlertTypes();
    }
  }

  /**
   * Save alert types to SharePoint
   */
  public async saveAlertTypes(alertTypes: IAlertType[]): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      // Clear existing items
      const alertTypesListApi = await this.getAlertTypesListApi(siteId);

      const existingItems = await this.graphClient
        .api(`${alertTypesListApi}/items`)
        .expand('fields')
        .get();

      for (const item of existingItems.value) {
        await this.graphClient
          .api(`${alertTypesListApi}/items/${item.id}`)
          .delete();
      }

      // Add new items
      for (let i = 0; i < alertTypes.length; i++) {
        const alertType = alertTypes[i];
        const listItem = {
          fields: {
            Title: alertType.name,
            IconName: alertType.iconName,
            BackgroundColor: alertType.backgroundColor,
            TextColor: alertType.textColor,
            AdditionalStyles: alertType.additionalStyles || '',
            PriorityStyles: JsonUtils.safeStringify(alertType.priorityStyles || {}) || '{}',
            SortOrder: i
          }
        };

        await this.graphClient
          .api(`${alertTypesListApi}/items`)
          .post(listItem);
      }
    } catch (error) {
      // Enhanced error handling for permission and access issues
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        logger.warn('SharePointAlertService', 'Access denied when trying to save alert types to SharePoint. Changes will be stored locally only');
        throw new Error('PERMISSION_DENIED: Cannot save alert types to SharePoint due to insufficient permissions. Changes stored locally only.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        logger.warn('SharePointAlertService', 'SharePoint alert types list not found. Cannot save alert types');
        throw new Error('LISTS_NOT_FOUND: SharePoint alert types list does not exist. Cannot save changes.');
      } else {
        logger.error('SharePointAlertService', 'Failed to save alert types', error);
        throw new Error(`SAVE_ALERT_TYPES_FAILED: ${error.message || 'Unknown error when saving alert types'}`);
      }
    }
  }

  /**
   * Map SharePoint list item to alert object using consolidated transformer
   */
  private mapSharePointItemToAlert(item: any, siteId?: string): IAlertItem {
    const fields = item.fields;

    // Debug log the raw SharePoint item to see what we're getting
    logger.debug('SharePointAlertService', 'Mapping SharePoint item to alert', {
      itemId: item.id,
      fieldKeys: Object.keys(fields),
      title: fields.Title,
      description: fields.Description,
      alertType: fields.AlertType,
      rawFields: fields
    });

    // Use AlertTransformers with _originalListItem included for multi-language support
    return AlertTransformers.mapSharePointItemToAlert(
      item,
      siteId || item.id.toString(),
      true // Include _originalListItem for SharePointAlertService
    );
  }

  /**
   * Repair the alerts list by removing outdated fields and adding current ones
   */
  public async repairAlertsList(siteId: string, progressCallback?: (message: string, progress: number) => void): Promise<IRepairResult> {
    logger.info('SharePointAlertService', `Starting repair of alerts list for site: ${siteId}`);
    
    const result: IRepairResult = {
      success: false,
      message: '',
      details: {
        columnsRemoved: [],
        columnsAdded: [],
        columnsUpdated: [],
        errors: [],
        warnings: []
      }
    };

    try {
      progressCallback?.('Analyzing current list structure...', 10);

      // First, verify the list exists and we have access
      let alertsListApi: string;
      let alertsList;
      try {
        alertsListApi = await this.getAlertsListApi(siteId);
        alertsList = await this.graphClient
          .api(alertsListApi)
          .get();
      } catch (error) {
        throw new Error(`Cannot access alerts list: ${error.message}. Please ensure you have proper permissions.`);
      }

      progressCallback?.('Retrieving current column information...', 20);

      // Get current list columns
      const currentColumns = await this.graphClient
        .api(`${alertsListApi}/columns`)
        .get();

      // Get all non-system columns that might be outdated
      const customColumns = currentColumns.value.filter((col: any) => 
        !col.readOnly && 
        col.name !== 'Title' && 
        col.name !== 'Created' && 
        col.name !== 'Modified' &&
        col.name !== 'Author' &&
        col.name !== 'Editor' &&
        col.name !== 'ID' &&
        !col.name.startsWith('_') // Exclude system columns
      );

      logger.info('SharePointAlertService', `Found ${customColumns.length} custom columns to evaluate`);
      progressCallback?.(`Found ${customColumns.length} custom columns to evaluate...`, 30);

      // Define current schema columns
      const keepColumns = [
        'Title', 'Description', 'AlertType', 'Priority', 'IsPinned', 
        'NotificationType', 'LinkUrl', 'LinkDescription', 'TargetSites', 
        'Status', 'ScheduledStart', 'ScheduledEnd', 'Metadata',
        'ItemType', 'TargetLanguage', 'LanguageGroup', 'AvailableForAll',
        'TargetUsers'
      ];

      // Language-specific columns are no longer needed - we use separate items per language

      progressCallback?.('Removing outdated columns...', 40);

      // Remove outdated custom columns
      let removedCount = 0;
      for (const column of customColumns) {
        if (!keepColumns.includes(column.name)) {
          try {
            await this.graphClient
              .api(`${alertsListApi}/columns/${column.id}`)
              .delete();
            
            result.details.columnsRemoved.push(column.name);
            removedCount++;
            logger.info('SharePointAlertService', `Removed outdated column: ${column.name}`);
            
            progressCallback?.(`Removed column: ${column.name}`, 40 + (removedCount * 20 / Math.max(customColumns.length, 1)));
          } catch (error) {
            const errorMsg = `Could not remove column ${column.name}: ${error.message}`;
            result.details.warnings.push(errorMsg);
            logger.warn('SharePointAlertService', errorMsg);
          }
        }
      }

      progressCallback?.('Adding/updating current columns...', 70);

      // Add current columns with updated definitions
      try {
        await this.addAlertsListColumns(siteId);
        
        // Get the expected columns that should have been added
        const expectedColumns = this.getExpectedAlertListColumns();
        result.details.columnsAdded = expectedColumns.map(col => col.name);
        
        progressCallback?.('Validating column structure...', 85);
      } catch (error) {
        const errorMsg = `Failed to add current columns: ${error.message}`;
        result.details.errors.push(errorMsg);
        logger.error('SharePointAlertService', errorMsg);
      }

      // Final validation - check if all expected columns exist
      progressCallback?.('Performing final validation...', 90);
      
      try {
        const finalColumns = await this.graphClient
          .api(`${alertsListApi}/columns`)
          .get();
        
        const finalColumnNames = finalColumns.value.map((col: any) => col.name);
        const missingColumns = keepColumns.filter(colName => !finalColumnNames.includes(colName));
        
        if (missingColumns.length > 0) {
          result.details.warnings.push(`Some expected columns are still missing: ${missingColumns.join(', ')}`);
        }
      } catch (error) {
        result.details.warnings.push(`Could not validate final column structure: ${error.message}`);
      }

      progressCallback?.('Repair completed successfully!', 100);

      const hasErrors = result.details.errors.length > 0;
      const hasWarnings = result.details.warnings.length > 0;
      
      result.success = !hasErrors;
      
      if (hasErrors) {
        result.message = `Repair completed with ${result.details.errors.length} error(s)`;
      } else if (hasWarnings) {
        result.message = `Repair completed successfully with ${result.details.warnings.length} warning(s)`;
      } else {
        result.message = 'Alerts list repair completed successfully';
      }
      
      result.message += `. Removed ${result.details.columnsRemoved.length} outdated column(s), added/updated ${result.details.columnsAdded.length} current column(s).`;

      logger.info('SharePointAlertService', result.message);
      return result;

    } catch (error) {
      const errorMessage = `Failed to repair alerts list: ${error.message}`;
      result.details.errors.push(errorMessage);
      result.message = errorMessage;
      logger.error('SharePointAlertService', errorMessage, error);
      return result;
    }
  }

  /**
   * Get the current site ID from context
   */
  public getCurrentSiteId(): string {
    return this.context.pageContext.site.id.toString();
  }

  /**
   * Get expected column definitions for validation
   */
  private getExpectedAlertListColumns(): any[] {
    return [
      { name: 'AlertType' },
      { name: 'Priority' },
      { name: 'IsPinned' },
      { name: 'NotificationType' },
      { name: 'LinkUrl' },
      { name: 'LinkDescription' },
      { name: 'TargetSites' },
      { name: 'Status' },
      { name: 'ScheduledStart' },
      { name: 'ScheduledEnd' },
      { name: 'Metadata' },
      { name: 'Description' },
      { name: 'ItemType' },
      { name: 'TargetLanguage' },
      { name: 'LanguageGroup' },
      { name: 'AvailableForAll' },
      { name: 'TargetUsers' }
    ];
  }

  /**
   * Map SharePoint list item to alert type object
   */
  private mapSharePointItemToAlertType(item: any): IAlertType {
    const fields = item.fields;
    return {
      name: fields.Title || '',
      iconName: fields.IconName || 'Info',
      backgroundColor: fields.BackgroundColor || '#0078d4',
      textColor: fields.TextColor || '#ffffff',
      additionalStyles: fields.AdditionalStyles || '',
      priorityStyles: JsonUtils.safeParse(fields.PriorityStyles) || {}
    };
  }

  /**
   * Get default alert types for fallback
   */
  private getDefaultAlertTypes(): IAlertType[] {
    return [
      {
        name: "Info",
        iconName: "Info",
        backgroundColor: "#389899",
        textColor: "#ffffff",
        additionalStyles: "",
        priorityStyles: {
          [AlertPriority.Critical]: "border: 2px solid #E81123;",
          [AlertPriority.High]: "border: 1px solid #EA4300;",
          [AlertPriority.Medium]: "",
          [AlertPriority.Low]: ""
        }
      },
      {
        name: "Warning",
        iconName: "Warning",
        backgroundColor: "#f1c40f",
        textColor: "#000000",
        additionalStyles: "",
        priorityStyles: {
          [AlertPriority.Critical]: "border: 2px solid #E81123;",
          [AlertPriority.High]: "border: 1px solid #EA4300;",
          [AlertPriority.Medium]: "",
          [AlertPriority.Low]: ""
        }
      },
      {
        name: "Maintenance",
        iconName: "ConstructionCone",
        backgroundColor: "#afd6d6",
        textColor: "#000000",
        additionalStyles: "",
        priorityStyles: {
          [AlertPriority.Critical]: "border: 2px solid #E81123;",
          [AlertPriority.High]: "border: 1px solid #EA4300;",
          [AlertPriority.Medium]: "",
          [AlertPriority.Low]: ""
        }
      },
      {
        name: "Interruption",
        iconName: "Error",
        backgroundColor: "#c54644",
        textColor: "#ffffff",
        additionalStyles: "",
        priorityStyles: {
          [AlertPriority.Critical]: "border: 2px solid #E81123;",
          [AlertPriority.High]: "border: 1px solid #EA4300;",
          [AlertPriority.Medium]: "",
          [AlertPriority.Low]: ""
        }
      }
    ];
  }

  /**
   * Get active alerts for display (considering scheduling)
   */
  public async getActiveAlerts(siteIds?: string[]): Promise<IAlertItem[]> {
    const allAlerts = await this.getAlerts(siteIds);
    const now = new Date();

    return allAlerts.filter(alert => {
      // Check if alert is scheduled and within active period
      // If scheduledStart exists and is in the future, not yet active
      if (alert.scheduledStart && new Date(alert.scheduledStart) > now) {
        return false; // Not yet active
      }

      // If scheduledEnd exists and is in the past, already expired
      if (alert.scheduledEnd && new Date(alert.scheduledEnd) < now) {
        return false; // Already expired
      }

      // Alert is active if:
      // 1. Status is 'Active' (regardless of dates)
      // 2. Status is 'Scheduled' and start time has passed (or no start time = forever)
      // 3. No dates at all means it's a forever alert
      return alert.status === 'Active' ||
        (alert.status === 'Scheduled' &&
          (!alert.scheduledStart || new Date(alert.scheduledStart) <= now));
    });
  }


  /**
   * Update alert status based on scheduling
   */
  public async updateAlertStatuses(): Promise<void> {
    try {
      const allAlerts = await this.getAlerts();
      const now = new Date();
      const updatesNeeded: { id: string, status: string }[] = [];

      for (const alert of allAlerts) {
        let newStatus = alert.status;

        if (alert.scheduledEnd && new Date(alert.scheduledEnd) < now && alert.status !== 'Expired') {
          newStatus = 'Expired';
        } else if (alert.scheduledStart && new Date(alert.scheduledStart) <= now && alert.status === 'Scheduled') {
          newStatus = 'Active';
        }

        if (newStatus !== alert.status) {
          updatesNeeded.push({ id: alert.id, status: newStatus });
        }
      }

      // Batch update statuses
      for (const update of updatesNeeded) {
        await this.updateAlert(update.id, { status: update.status as any });
      }
    } catch (error) {
      logger.error('SharePointAlertService', 'Failed to update alert statuses', error);
    }
  }


  /**
   * Get supported languages from TargetLanguage choice field
   */
  public async getSupportedLanguages(): Promise<string[]> {
    try {
      // Use SharePoint REST API for consistency with updateTargetLanguageChoices
      const webAbsoluteUrl = this.context.pageContext.web.absoluteUrl;
      const fieldInfoUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${this.alertsListName}')/fields/getbytitle('TargetLanguage')`;
      
      const fieldResponse = await this.context.spHttpClient.get(fieldInfoUrl, 
        SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata.metadata=minimal'
          }
        });
      
      if (fieldResponse.ok) {
        const fieldData = await fieldResponse.json();
        const targetLanguageColumn = fieldData.value || fieldData;
        const choices = targetLanguageColumn.Choices || ['en-us'];
        
        // Filter out 'all' and return actual language codes
        return choices.filter((choice: string) => choice !== 'all');
      }

      return ['en-us']; // Default fallback
    } catch (error) {
      logger.warn('SharePointAlertService', 'Failed to get supported languages:', error);
      return ['en-us'];
    }
  }

  /**
   * Add a language to the TargetLanguage choice field
   */
  public async addLanguageSupport(languageCode: string): Promise<void> {
    try {
      await this.updateTargetLanguageChoices('add', languageCode);
      logger.info('SharePointAlertService', `Successfully added language support for ${languageCode}`);
    } catch (error) {
      logger.error('SharePointAlertService', `Error adding language support for ${languageCode}:`, error);
      throw error;
    }
  }

  /**
   * Remove a language from the TargetLanguage choice field
   */
  public async removeLanguageSupport(languageCode: string): Promise<void> {
    try {
      await this.updateTargetLanguageChoices('remove', languageCode);
      logger.info('SharePointAlertService', `Successfully removed language support for ${languageCode}`);
    } catch (error) {
      logger.error('SharePointAlertService', `Error removing language support for ${languageCode}:`, error);
      throw error;
    }
  }

  /**
   * Update the TargetLanguage choice field choices
   */
  private async updateTargetLanguageChoices(action: 'add' | 'remove', languageCode: string): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      // Get current TargetLanguage column using SharePoint REST API for consistency
      const webAbsoluteUrl = this.context.pageContext.web.absoluteUrl;
      const fieldInfoUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${this.alertsListName}')/fields/getbytitle('TargetLanguage')`;
      
      const fieldResponse = await this.context.spHttpClient.get(fieldInfoUrl, 
        SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata.metadata=minimal'
          }
        });
      
      if (!fieldResponse.ok) {
        logger.warn('SharePointAlertService', 'TargetLanguage column not found via REST API');
        return;
      }
      
      const fieldData = await fieldResponse.json();
      const targetLanguageColumn = fieldData.value || fieldData;
      const currentChoices = targetLanguageColumn.Choices || ['all', 'en-us'];

      logger.info('SharePointAlertService', `Current TargetLanguage choices from REST API:`, { currentChoices });

      let updatedChoices: string[];
      if (action === 'add') {
        // Add the language if not already present
        if (!currentChoices.includes(languageCode)) {
          updatedChoices = [...currentChoices, languageCode].sort();
        } else {
          updatedChoices = currentChoices;
          logger.info('SharePointAlertService', `Language ${languageCode} already exists in choices`);
          return; // No update needed
        }
      } else {
        // Remove the language (but keep 'all' and 'en-us')
        updatedChoices = currentChoices.filter((choice: string) => 
          choice !== languageCode || choice === 'all' || choice === 'en-us'
        );
        if (updatedChoices.length === currentChoices.length) {
          logger.info('SharePointAlertService', `Language ${languageCode} not found in choices`);
          return; // No update needed
        }
      }

      logger.info('SharePointAlertService', `Updating TargetLanguage choices from [${currentChoices.join(', ')}] to [${updatedChoices.join(', ')}]`);

      // Use SharePoint REST API (the correct approach for choice field schema updates)
      // Graph API cannot update choice field schemas, only REST API works
      logger.info('SharePointAlertService', 'Using SharePoint REST API approach for choice field schema update');
      
      // Prepare the update payload for REST API (OData v4.0 format)
      const updatePayload = {
        '@odata.type': 'SP.FieldChoice',
        Choices: updatedChoices
      };
      
      // Update the field via REST API
      const updateUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${this.alertsListName}')/fields/getbytitle('TargetLanguage')`;
      
      const updateResponse = await this.context.spHttpClient.post(updateUrl, 
        SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata.metadata=minimal',
            'Content-Type': 'application/json;odata.metadata=minimal',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': targetLanguageColumn['@odata.etag'] || targetLanguageColumn.etag || '*'
          },
          body: JSON.stringify(updatePayload)
        });
      
      if (!updateResponse.ok && updateResponse.status !== 204) {
        const errorText = await updateResponse.text();
        throw new Error(`REST API update failed: ${updateResponse.status} ${updateResponse.statusText} - ${errorText}`);
      }

      logger.info('SharePointAlertService', `Successfully updated TargetLanguage choices:`, { 
        action, 
        languageCode, 
        updatedChoices 
      });

    } catch (error) {
      logger.error('SharePointAlertService', 'Failed to update TargetLanguage choices:', error);
      
      // More detailed error information
      if (error.code === 'BadRequest') {
        logger.error('SharePointAlertService', 'BadRequest details:', {
          message: error.message,
          requestId: error['request-id'],
          correlationId: error['correlation-id']
        });
      }
      
      throw new Error(`Failed to update TargetLanguage choices: ${error.message || error}`);
    }
  }
}
