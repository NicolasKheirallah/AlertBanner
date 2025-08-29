import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { AlertPriority, NotificationType, IAlertType, ITargetingRule, IAlertRichMedia, IQuickAction } from "../Alerts/IAlerts";

export interface IAlertItem {
  id: string;
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  targetingRules?: ITargetingRule[];
  notificationType: NotificationType;
  richMedia?: IAlertRichMedia;
  linkUrl?: string;
  linkDescription?: string;
  quickActions?: IQuickAction[];
  targetSites: string[];
  status: 'Active' | 'Expired' | 'Scheduled';
  createdDate: string;
  createdBy: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  metadata?: any;
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

  // Dynamic language support - for additional languages
  [key: string]: any;
}

export class SharePointAlertService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private alertsListName = 'Alerts';
  private alertTypesListName = 'AlertBannerTypes';

  constructor(graphClient: MSGraphClientV3, context: ApplicationCustomizerContext) {
    this.graphClient = graphClient;
    this.context = context;
  }

  /**
   * Initialize SharePoint lists if they don't exist
   */
  /**
   * Check which sites need list creation
   */
  public async checkListsNeeded(): Promise<{ site: string; needsAlerts: boolean; needsTypes: boolean }[]> {
    const results = [];
    const currentSiteId = this.context.pageContext.site.id.toString();
    
    // Check current site
    let needsAlerts = false;
    let needsTypes = false;
    
    try {
      await this.graphClient.api(`/sites/${currentSiteId}/lists/Alerts`).get();
    } catch (error) {
      if (error.message?.includes('not found') || error.message?.includes('404')) {
        needsAlerts = true;
      }
    }
    
    try {
      await this.graphClient.api(`/sites/${currentSiteId}/lists/AlertBannerTypes`).get();
    } catch (error) {
      if (error.message?.includes('not found') || error.message?.includes('404')) {
        needsTypes = true;
      }
    }
    
    results.push({
      site: currentSiteId,
      needsAlerts,
      needsTypes
    });
    
    return results;
  }

  public async initializeLists(): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      // Check if alerts list exists or can be created
      let alertsListCreated = false;
      let typesListCreated = false;
      
      try {
        alertsListCreated = await this.ensureAlertsList(siteId);
        if (alertsListCreated) {
          console.log('Alert Banner alerts list created successfully');
        } else {
          console.log('Alert Banner alerts list already exists');
        }
      } catch (alertsError) {
        if (alertsError.message?.includes('PERMISSION_DENIED')) {
          console.warn('Cannot create alerts list due to insufficient permissions. Alert functionality may be limited.');
          // Don't throw here, continue with types list
        } else {
          throw alertsError;
        }
      }

      try {
        typesListCreated = await this.ensureAlertTypesList(siteId);
        if (typesListCreated) {
          console.log('Alert Banner types list created successfully');
        } else {
          console.log('Alert Banner types list already exists');
        }
      } catch (typesError) {
        if (typesError.message?.includes('PERMISSION_DENIED')) {
          console.warn('Cannot create alert types list due to insufficient permissions. Default alert types will be used.');
          // Don't throw here, app can still function with default types
        } else {
          throw typesError;
        }
      }
    } catch (error) {
      // Enhanced error handling for common permission issues
      if (error.message?.includes('PERMISSION_DENIED')) {
        console.warn('SharePoint list creation failed due to insufficient permissions.');
        throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        console.warn('SharePoint lists not found and cannot be created.');
        throw new Error('LISTS_NOT_FOUND: SharePoint lists do not exist and cannot be created.');
      } else {
        console.error('Failed to initialize SharePoint lists:', error);
        throw new Error(`INITIALIZATION_FAILED: ${error.message || 'Unknown error during SharePoint initialization'}`);
      }
    }
  }

  /**
   * Create alerts list if it doesn't exist
   */
  private async ensureAlertsList(siteId: string): Promise<boolean> {
    try {
      // Try to get the list first
      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}`)
        .get();
      return false; // List already exists
    } catch (error) {
      // Check if it's a permission error or list doesn't exist
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        console.warn(`Cannot access or create alerts list due to insufficient permissions.`);
        throw new Error('PERMISSION_DENIED: User lacks permissions to access or create SharePoint lists.');
      }

      // Check if user has permission to create lists before attempting
      try {
        // Test permissions by trying to get all lists
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .select('id')
          .top(1)
          .get();
      } catch (permissionError) {
        const errorMessage = permissionError.message || '';
        const statusCode = permissionError.code || '';
        
        console.error('Permission check failed:', {
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

      // List doesn't exist, create it
      console.log('Creating alerts list...');

      const listDefinition = {
        displayName: this.alertsListName,
        list: {
          template: 'genericList'
        }
      };

      try {
        console.log('📋 Creating basic list structure...');
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .post(listDefinition);
        console.log('✅ Basic list structure created successfully');

        // Add custom columns after list creation
        console.log('🏗️ Adding custom columns to Alerts list...');
        await this.addAlertsListColumns(siteId);
        console.log('✅ All custom columns added successfully');

        return true; // List was created
      } catch (createError) {
        if (createError.message?.includes('Access denied') || createError.message?.includes('403')) {
          console.warn('User lacks permissions to create SharePoint lists.');
          throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
        }
        if (createError.message?.includes('CRITICAL_COLUMNS_FAILED')) {
          console.error('List created but critical columns failed:', createError.message);
          throw new Error(`LIST_INCOMPLETE: ${createError.message}`);
        }
        throw createError;
      }
    }
  }

  /**
   * Add custom columns to the Alerts list after creation
   */
  private async addAlertsListColumns(siteId: string): Promise<void> {
    // Get the AlertBannerTypes list ID for the lookup field
    let alertTypesListId = '';
    try {
      const alertTypesList = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertTypesListName}`)
        .select('id')
        .get();
      alertTypesListId = alertTypesList.id;
    } catch (error) {
      console.warn('Could not get AlertBannerTypes list ID for lookup field:', error);
      // If we can't get the list ID, we'll create AlertType as a text field instead
    }

    const columns = [
      // Create AlertType as lookup if we have the AlertBannerTypes list, otherwise as text
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
          maxLength: 255,
          allowMultipleLines: false
        }
      },
      {
        name: 'Priority',
        text: { 
          maxLength: 50,
          allowMultipleLines: false
        }
      },
      {
        name: 'IsPinned',
        boolean: {}
      },
      {
        name: 'NotificationType',
        text: { 
          maxLength: 50,
          allowMultipleLines: false
        }
      },
      {
        name: 'LinkUrl',
        text: { 
          maxLength: 2083,
          allowMultipleLines: false
        }
      },
      {
        name: 'LinkDescription',
        text: { 
          maxLength: 255,
          allowMultipleLines: false
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
        text: { 
          maxLength: 50,
          allowMultipleLines: false
        }
      },
      {
        name: 'ScheduledStart',
        dateTime: {
          displayFormat: 'friendlyDateTime'
        },
        indexed: true
      },
      {
        name: 'ScheduledEnd',
        dateTime: {
          displayFormat: 'friendlyDateTime'
        },
        indexed: true
      },
      {
        name: 'Metadata',
        text: { 
          allowMultipleLines: true,
          maxLength: 4000
        }
      },
      {
        name: 'Description',
        text: { 
          allowMultipleLines: true,
          maxLength: 4000,
          textType: 'richText',
          linesForEditing: 10,
          appendChangesToExistingText: false
        }
      }
    ];

    const criticalColumns = ['ScheduledStart', 'ScheduledEnd'];
    const failedColumns: string[] = [];

    for (const column of columns) {
      try {
        console.log(`Creating column: ${column.name}`);
        await this.graphClient
          .api(`/sites/${siteId}/lists/${this.alertsListName}/columns`)
          .post(column);
        console.log(`✅ Successfully created column: ${column.name}`);
      } catch (error) {
        console.error(`❌ Failed to create column ${column.name}:`, {
          columnName: column.name,
          columnDefinition: column,
          error: error.message,
          statusCode: error.code || error.status
        });
        
        failedColumns.push(column.name);
        
        // For critical columns, try alternative creation methods
        if (criticalColumns.includes(column.name) && column.name.includes('Scheduled')) {
          console.log(`🔄 Retrying critical column ${column.name} with alternative method...`);
          try {
            // Try with simpler dateTime definition
            const simpleColumn = {
              name: column.name,
              dateTime: {}
            };
            await this.graphClient
              .api(`/sites/${siteId}/lists/${this.alertsListName}/columns`)
              .post(simpleColumn);
            console.log(`✅ Successfully created ${column.name} with alternative method`);
          } catch (retryError) {
            console.error(`❌ Alternative method also failed for ${column.name}:`, retryError);
          }
        }
      }
    }

    // Report summary of column creation
    if (failedColumns.length > 0) {
      console.warn(`⚠️ Column creation summary: ${failedColumns.length} columns failed:`, failedColumns);
      
      // If critical columns failed, throw an error
      const failedCriticalColumns = failedColumns.filter(name => criticalColumns.includes(name));
      if (failedCriticalColumns.length > 0) {
        console.error(`🚨 Critical columns failed: ${failedCriticalColumns.join(', ')}`);
        throw new Error(`CRITICAL_COLUMNS_FAILED: Failed to create critical columns: ${failedCriticalColumns.join(', ')}`);
      }
    } else {
      console.log(`✅ All ${columns.length} columns created successfully`);
    }
  }

  /**
   * Create alert types list if it doesn't exist
   */
  private async ensureAlertTypesList(siteId: string): Promise<boolean> {
    try {
      // Try to get the list first
      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertTypesListName}`)
        .get();
      return false; // List already exists
    } catch (error) {
      // Check if it's a permission error or list doesn't exist
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        console.warn(`Cannot access or create alert types list due to insufficient permissions.`);
        throw new Error('PERMISSION_DENIED: User lacks permissions to access or create SharePoint lists.');
      }

      // Check if user has permission to create lists before attempting
      try {
        // Test permissions by trying to get all lists
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .select('id')
          .top(1)
          .get();
      } catch (permissionError) {
        if (permissionError.message?.includes('Access denied') || permissionError.message?.includes('403')) {
          console.warn('User lacks permissions to create SharePoint lists.');
          throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
        }
      }

      // List doesn't exist, create it
      console.log('Creating alert types list...');

      const listDefinition = {
        displayName: this.alertTypesListName,
        list: {
          template: 'genericList'
        }
      };

      try {
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .post(listDefinition);

        // Add custom columns after list creation
        await this.addAlertTypesListColumns(siteId);

        return true; // List was created
      } catch (createError) {
        if (createError.message?.includes('Access denied') || createError.message?.includes('403')) {
          console.warn('User lacks permissions to create SharePoint lists.');
          throw new Error('PERMISSION_DENIED: User lacks permissions to create SharePoint lists.');
        }
        throw createError;
      }
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

    for (const column of columns) {
      try {
        await this.graphClient
          .api(`/sites/${siteId}/lists/${this.alertTypesListName}/columns`)
          .post(column);
      } catch (error) {
        console.warn(`Failed to create AlertTypes column ${column.name}:`, error);
        // Continue creating other columns even if one fails
      }
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
          console.warn('Failed to get hierarchical sites, falling back to current site:', error);
          sitesToQuery = [this.context.pageContext.site.id.toString()];
        }
      }
      const allAlerts: IAlertItem[] = [];
      const seenAlerts = new Map<string, IAlertItem>(); // Track alerts by title+description to avoid duplicates

      // Query alerts from each site
      for (const siteId of sitesToQuery) {
        try {
          const response = await this.graphClient
            .api(`/sites/${siteId}/lists/${this.alertsListName}/items`)
            .expand('fields')
            .get();

          const siteAlerts = response.value.map((item: any) => this.mapSharePointItemToAlert(item, siteId));
          
          // Deduplicate alerts based on SharePoint item ID and site ID
          for (const alert of siteAlerts) {
            const dedupeKey = `${siteId}-${alert.id.split('-').pop()}`; // Use actual SharePoint item ID
            if (!seenAlerts.has(dedupeKey)) {
              seenAlerts.set(dedupeKey, alert);
              allAlerts.push(alert);
            } else {
              console.warn(`Duplicate alert detected and skipped: ${alert.title} (ID: ${alert.id})`);
            }
          }
        } catch (error) {
          console.warn(`Failed to get alerts from site ${siteId}:`, error);
          // Continue with other sites
        }
      }

      return allAlerts.sort((a, b) => new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime());
    } catch (error) {
      // Enhanced error handling for permission and access issues
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        console.warn('Access denied when trying to get alerts from SharePoint.');
        throw new Error('PERMISSION_DENIED: Cannot access SharePoint alerts due to insufficient permissions.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        console.warn('SharePoint alerts list not found.');
        throw new Error('LISTS_NOT_FOUND: SharePoint alerts list does not exist.');
      } else {
        console.error('Failed to get alerts:', error);
        throw new Error(`GET_ALERTS_FAILED: ${error.message || 'Unknown error when retrieving alerts'}`);
      }
    }
  }

  /**
   * Create a new alert
   */
  public async createAlert(alert: Omit<IAlertItem, 'id' | 'createdDate' | 'createdBy' | 'status'>): Promise<IAlertItem> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      const listItem = {
        fields: {
          Title: alert.title,
          Description: alert.description,
          AlertType: alert.AlertType, // This should be the lookup value (just the text name)
          Priority: alert.priority,
          IsPinned: alert.isPinned,
          NotificationType: alert.notificationType,
          LinkUrl: alert.linkUrl || '',
          LinkDescription: alert.linkDescription || '',
          TargetSites: JSON.stringify(alert.targetSites),
          Status: alert.scheduledStart && new Date(alert.scheduledStart) > new Date() ? 'Scheduled' : 'Active',
          ScheduledStart: alert.scheduledStart || null,
          ScheduledEnd: alert.scheduledEnd || null,
          Metadata: alert.metadata ? JSON.stringify(alert.metadata) : ''
        }
      };

      const response = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items`)
        .post(listItem);

      // Get the created item with expanded fields
      const createdItem = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${response.id}`)
        .expand('fields')
        .get();

      return this.mapSharePointItemToAlert(createdItem, siteId);
    } catch (error) {
      console.error('Failed to create alert:', error);
      throw error;
    }
  }

  /**
   * Extract site ID and item ID from composite alert ID
   */
  private parseAlertId(alertId: string): { siteId: string; itemId: string } {
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

  /**
   * Update an existing alert
   */
  public async updateAlert(alertId: string, updates: Partial<IAlertItem>): Promise<IAlertItem> {
    try {
      const { siteId, itemId } = this.parseAlertId(alertId);

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
          ...(updates.targetSites && { TargetSites: JSON.stringify(updates.targetSites) }),
          ...(updates.scheduledStart !== undefined && { ScheduledStart: updates.scheduledStart }),
          ...(updates.scheduledEnd !== undefined && { ScheduledEnd: updates.scheduledEnd }),
          ...(updates.metadata && { Metadata: JSON.stringify(updates.metadata) })
        }
      };

      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${itemId}/fields`)
        .patch(listItem.fields);

      // Get the updated item
      const updatedItem = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${itemId}`)
        .expand('fields')
        .get();

      return this.mapSharePointItemToAlert(updatedItem, siteId);
    } catch (error) {
      console.error('Failed to update alert:', error);
      throw error;
    }
  }

  /**
   * Delete an alert
   */
  public async deleteAlert(alertId: string): Promise<void> {
    try {
      const { siteId, itemId } = this.parseAlertId(alertId);

      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${itemId}`)
        .delete();
    } catch (error) {
      console.error('Failed to delete alert:', error);
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
  public async getAlertTypes(): Promise<IAlertType[]> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

      // Try to ensure the alert types list exists
      try {
        await this.ensureAlertTypesList(siteId);
      } catch (ensureError) {
        console.warn('Could not ensure alert types list exists:', ensureError);
      }

      const response = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertTypesListName}/items`)
        .expand('fields')
        .orderby('fields/SortOrder')
        .get();

      return response.value.map((item: any) => this.mapSharePointItemToAlertType(item));
    } catch (error) {
      console.warn('Failed to get alert types from SharePoint, using defaults:', error);
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
      const existingItems = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertTypesListName}/items`)
        .expand('fields')
        .get();

      for (const item of existingItems.value) {
        await this.graphClient
          .api(`/sites/${siteId}/lists/${this.alertTypesListName}/items/${item.id}`)
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
            PriorityStyles: JSON.stringify(alertType.priorityStyles || {}),
            SortOrder: i
          }
        };

        await this.graphClient
          .api(`/sites/${siteId}/lists/${this.alertTypesListName}/items`)
          .post(listItem);
      }
    } catch (error) {
      // Enhanced error handling for permission and access issues
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
        console.warn('Access denied when trying to save alert types to SharePoint. Changes will be stored locally only.');
        throw new Error('PERMISSION_DENIED: Cannot save alert types to SharePoint due to insufficient permissions. Changes stored locally only.');
      } else if (error.message?.includes('404') || error.message?.includes('not found')) {
        console.warn('SharePoint alert types list not found. Cannot save alert types.');
        throw new Error('LISTS_NOT_FOUND: SharePoint alert types list does not exist. Cannot save changes.');
      } else {
        console.error('Failed to save alert types:', error);
        throw new Error(`SAVE_ALERT_TYPES_FAILED: ${error.message || 'Unknown error when saving alert types'}`);
      }
    }
  }

  /**
   * Map SharePoint list item to alert object
   */
  private mapSharePointItemToAlert(item: any, siteId?: string): IAlertItem {
    const fields = item.fields;
    
    // Create the original list item for multi-language support
    const originalListItem: IAlertListItem = {
      Id: parseInt(item.id.toString()),
      Title: fields.Title || '',
      Description: fields.Description || '',
      AlertType: fields.AlertType?.LookupValue || fields.AlertType || '',
      Priority: fields.Priority || AlertPriority.Medium,
      IsPinned: fields.IsPinned || false,
      NotificationType: fields.NotificationType || NotificationType.None,
      LinkUrl: fields.LinkUrl || '',
      LinkDescription: fields.LinkDescription || '',
      TargetSites: fields.TargetSites || '',
      Status: fields.Status || 'Active',
      Created: fields.Created || item.createdDateTime,
      Author: {
        Title: item.createdBy?.user?.displayName || item.author?.Title || 'Unknown'
      },
      ScheduledStart: fields.ScheduledStart || undefined,
      ScheduledEnd: fields.ScheduledEnd || undefined,
      Metadata: fields.Metadata || undefined,
      
      // Add all multi-language fields
      Title_EN: fields.Title_EN || '',
      Title_FR: fields.Title_FR || '',
      Title_DE: fields.Title_DE || '',
      Title_ES: fields.Title_ES || '',
      Title_SV: fields.Title_SV || '',
      Title_FI: fields.Title_FI || '',
      Title_DA: fields.Title_DA || '',
      Title_NO: fields.Title_NO || '',
      
      Description_EN: fields.Description_EN || '',
      Description_FR: fields.Description_FR || '',
      Description_DE: fields.Description_DE || '',
      Description_ES: fields.Description_ES || '',
      Description_SV: fields.Description_SV || '',
      Description_FI: fields.Description_FI || '',
      Description_DA: fields.Description_DA || '',
      Description_NO: fields.Description_NO || '',
      
      LinkDescription_EN: fields.LinkDescription_EN || '',
      LinkDescription_FR: fields.LinkDescription_FR || '',
      LinkDescription_DE: fields.LinkDescription_DE || '',
      LinkDescription_ES: fields.LinkDescription_ES || '',
      LinkDescription_SV: fields.LinkDescription_SV || '',
      LinkDescription_FI: fields.LinkDescription_FI || '',
      LinkDescription_DA: fields.LinkDescription_DA || '',
      LinkDescription_NO: fields.LinkDescription_NO || '',
      
      // Include any additional dynamic language fields
      ...Object.keys(fields)
        .filter(key => key.match(/^(Title|Description|LinkDescription)_[A-Z]{2}$/))
        .reduce((acc, key) => ({ ...acc, [key]: fields[key] }), {})
    };

    return {
      id: siteId ? `${siteId}-${item.id}` : item.id.toString(),
      title: fields.Title || '',
      description: fields.Description || '',
      AlertType: fields.AlertType?.LookupValue || fields.AlertType || '',
      priority: fields.Priority || AlertPriority.Medium,
      isPinned: fields.IsPinned || false,
      notificationType: fields.NotificationType || NotificationType.None,
      linkUrl: fields.LinkUrl || '',
      linkDescription: fields.LinkDescription || '',
      targetSites: fields.TargetSites ? JSON.parse(fields.TargetSites) : [],
      status: fields.Status || 'Active',
      createdDate: fields.Created || item.createdDateTime,
      createdBy: item.createdBy?.user?.displayName || item.author?.Title || 'Unknown',
      scheduledStart: fields.ScheduledStart || undefined,
      scheduledEnd: fields.ScheduledEnd || undefined,
      metadata: fields.Metadata ? JSON.parse(fields.Metadata) : undefined,
      _originalListItem: originalListItem
    };
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
      priorityStyles: fields.PriorityStyles ? JSON.parse(fields.PriorityStyles) : {}
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
      console.error('Failed to update alert statuses:', error);
    }
  }

  /**
   * Get localized content for a specific field and language
   */
  public getLocalizedField(item: IAlertListItem, fieldName: string, languageCode: string): string {
    // Convert language code to uppercase format for field names (e.g., 'en-us' -> 'EN')
    const languageSuffix = languageCode.split('-')[0].toUpperCase();
    const localizedFieldName = `${fieldName}_${languageSuffix}`;

    // Try localized field first, then fall back to English, then original field
    return item[localizedFieldName] || 
           item[`${fieldName}_EN`] || 
           item[fieldName] || 
           '';
  }

  /**
   * Get all available languages for multi-language content
   */
  public getAvailableContentLanguages(item: IAlertListItem): string[] {
    const languages: string[] = [];
    const fieldPrefixes = ['Title_', 'Description_', 'LinkDescription_'];
    
    // Check which language fields have content
    Object.keys(item).forEach(key => {
      fieldPrefixes.forEach(prefix => {
        if (key.startsWith(prefix)) {
          const languageCode = key.substring(prefix.length).toLowerCase();
          const fullLanguageCode = this.mapLanguageCodeToFull(languageCode);
          if (item[key] && !languages.includes(fullLanguageCode)) {
            languages.push(fullLanguageCode);
          }
        }
      });
    });

    return languages;
  }

  /**
   * Map short language codes to full codes (e.g., 'EN' -> 'en-us')
   */
  private mapLanguageCodeToFull(shortCode: string): string {
    const languageMap: { [key: string]: string } = {
      'EN': 'en-us',
      'FR': 'fr-fr',
      'DE': 'de-de',
      'ES': 'es-es',
      'SV': 'sv-se',
      'FI': 'fi-fi',
      'DA': 'da-dk',
      'NO': 'nb-no'
    };

    return languageMap[shortCode.toUpperCase()] || shortCode.toLowerCase();
  }

  /**
   * Create or update list columns for additional languages dynamically
   */
  public async addLanguageColumns(languageCode: string): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const languageSuffix = languageCode.split('-')[0].toUpperCase();

      const columnsToAdd = [
        {
          name: `Title_${languageSuffix}`,
          text: { maxLength: 255 }
        },
        {
          name: `Description_${languageSuffix}`,
          text: { allowMultipleLines: true, appendChangesToExistingText: false }
        },
        {
          name: `LinkDescription_${languageSuffix}`,
          text: { maxLength: 255 }
        }
      ];

      for (const column of columnsToAdd) {
        try {
          await this.graphClient
            .api(`/sites/${siteId}/lists/${this.alertsListName}/columns`)
            .post(column);
          
          console.log(`Added column ${column.name} for language ${languageCode}`);
        } catch (error) {
          if (error.message?.includes('already exists')) {
            console.log(`Column ${column.name} already exists`);
          } else {
            console.warn(`Failed to add column ${column.name}:`, error);
          }
        }
      }
    } catch (error) {
      console.error(`Failed to add language columns for ${languageCode}:`, error);
      throw error;
    }
  }

  /**
   * Remove language columns from SharePoint lists
   */
  public async removeLanguageColumns(languageCode: string): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const languageSuffix = languageCode.split('-')[0].toUpperCase();

      const columnsToRemove = [
        `Title_${languageSuffix}`,
        `Description_${languageSuffix}`,
        `LinkDescription_${languageSuffix}`
      ];

      console.log(`🗑️ Removing language columns for ${languageCode} (${languageSuffix})...`);

      for (const columnName of columnsToRemove) {
        try {
          // First check if the column exists
          const existingColumns = await this.graphClient
            .api(`/sites/${siteId}/lists/${this.alertsListName}/columns`)
            .select('name,id')
            .filter(`name eq '${columnName}'`)
            .get();

          if (existingColumns.value && existingColumns.value.length > 0) {
            const columnId = existingColumns.value[0].id;
            await this.graphClient
              .api(`/sites/${siteId}/lists/${this.alertsListName}/columns/${columnId}`)
              .delete();
            
            console.log(`✅ Removed column ${columnName} for language ${languageCode}`);
          } else {
            console.log(`ℹ️ Column ${columnName} does not exist, skipping`);
          }
        } catch (error) {
          if (error.message?.includes('does not exist') || error.message?.includes('not found')) {
            console.log(`ℹ️ Column ${columnName} already does not exist`);
          } else {
            console.warn(`⚠️ Failed to remove column ${columnName}:`, error);
            // Don't throw - continue removing other columns
          }
        }
      }
      
      console.log(`✅ Completed removal process for language ${languageCode}`);
    } catch (error) {
      console.error(`❌ Failed to remove language columns for ${languageCode}:`, error);
      throw error;
    }
  }

  /**
   * Get localized content from an alert item
   */
  public getLocalizedAlertContent(alertItem: IAlertItem, languageCode: string): {
    title: string;
    description: string;
    linkDescription: string;
  } {
    if (!alertItem._originalListItem) {
      // Fallback to default fields if no original list item
      return {
        title: alertItem.title,
        description: alertItem.description,
        linkDescription: alertItem.linkDescription || ''
      };
    }

    return {
      title: this.getLocalizedField(alertItem._originalListItem, 'Title', languageCode),
      description: this.getLocalizedField(alertItem._originalListItem, 'Description', languageCode),
      linkDescription: this.getLocalizedField(alertItem._originalListItem, 'LinkDescription', languageCode)
    };
  }

  /**
   * Check if alert has content in specific language
   */
  public alertHasLanguageContent(alertItem: IAlertItem, languageCode: string): boolean {
    if (!alertItem._originalListItem) return false;

    const content = this.getLocalizedAlertContent(alertItem, languageCode);
    return !!(content.title || content.description || content.linkDescription);
  }

  /**
   * Get all languages that have content for a specific alert
   */
  public getAlertContentLanguages(alertItem: IAlertItem): string[] {
    if (!alertItem._originalListItem) return [];
    
    return this.getAvailableContentLanguages(alertItem._originalListItem);
  }

  /**
   * Update multi-language content for an alert
   */
  public async updateAlertMultiLanguageContent(
    alertId: string, 
    multiLanguageContent: { [languageCode: string]: { title?: string; description?: string; linkDescription?: string } }
  ): Promise<void> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const updateFields: any = {};

      // Convert multi-language content to SharePoint field format
      Object.entries(multiLanguageContent).forEach(([languageCode, content]) => {
        const languageSuffix = languageCode.split('-')[0].toUpperCase();
        
        if (content.title !== undefined) {
          updateFields[`Title_${languageSuffix}`] = content.title;
        }
        if (content.description !== undefined) {
          updateFields[`Description_${languageSuffix}`] = content.description;
        }
        if (content.linkDescription !== undefined) {
          updateFields[`LinkDescription_${languageSuffix}`] = content.linkDescription;
        }
      });

      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${alertId}/fields`)
        .patch(updateFields);

      console.log(`Updated multi-language content for alert ${alertId}`);
    } catch (error) {
      console.error('Failed to update multi-language content:', error);
      throw error;
    }
  }

  /**
   * Get supported language codes that have columns in the list
   */
  public async getSupportedLanguageColumns(): Promise<string[]> {
    try {
      const siteId = this.context.pageContext.site.id.toString();
      const response = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/columns`)
        .select('name')
        .get();

      const columns: any[] = response.value;
      const languageCodes = new Set<string>();

      columns.forEach(column => {
        const match = column.name.match(/^(Title|Description|LinkDescription)_([A-Z]{2})$/);
        if (match) {
          const languageCode = this.mapLanguageCodeToFull(match[2]);
          languageCodes.add(languageCode);
        }
      });

      return Array.from(languageCodes);
    } catch (error) {
      console.error('Failed to get supported language columns:', error);
      // Return default supported languages
      return ['en-us', 'fr-fr', 'de-de', 'es-es', 'sv-se', 'fi-fi', 'da-dk', 'nb-no'];
    }
  }
}