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
        description: 'Stores alert banner notifications',
        template: 'genericList',
        columns: [
          {
            name: 'Description',
            text: { allowMultipleLines: true, appendChangesToExistingText: false }
          },
          {
            name: 'AlertType',
            text: { maxLength: 255 }
          },
          {
            name: 'Priority',
            choice: {
              choices: ['Low', 'Medium', 'High', 'Critical'],
              displayAs: 'dropDownMenu',
              defaultValue: 'Medium'
            }
          },
          {
            name: 'IsPinned',
            boolean: {
              defaultValue: false
            }
          },
          {
            name: 'NotificationType',
            choice: {
              choices: ['None', 'Browser', 'Email', 'Both'],
              displayAs: 'dropDownMenu',
              defaultValue: 'None'
            }
          },
          {
            name: 'LinkUrl',
            text: { maxLength: 2083 }
          },
          {
            name: 'LinkDescription',
            text: { maxLength: 255 }
          },
          {
            name: 'TargetSites',
            text: { allowMultipleLines: true }
          },
          {
            name: 'Status',
            choice: {
              choices: ['Active', 'Expired', 'Scheduled'],
              displayAs: 'dropDownMenu',
              defaultValue: 'Active'
            }
          },
          {
            name: 'ScheduledStart',
            dateTime: { displayAs: 'default', format: 'dateTime' }
          },
          {
            name: 'ScheduledEnd',
            dateTime: { displayAs: 'default', format: 'dateTime' }
          },
          {
            name: 'Metadata',
            text: { allowMultipleLines: true }
          }
          // Note: Multi-language fields will be added dynamically based on user selection
        ]
      };

      try {
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .post(listDefinition);

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
        description: 'Stores alert banner type definitions',
        template: 'genericList',
        columns: [
          {
            name: 'IconName',
            text: { maxLength: 100 }
          },
          {
            name: 'BackgroundColor',
            text: { maxLength: 50 }
          },
          {
            name: 'TextColor',
            text: { maxLength: 50 }
          },
          {
            name: 'AdditionalStyles',
            text: { allowMultipleLines: true }
          },
          {
            name: 'PriorityStyles',
            text: { allowMultipleLines: true }
          },
          {
            name: 'SortOrder',
            number: { 
              decimalPlaces: 0,
              defaultValue: '0'
            }
          }
        ]
      };

      try {
        await this.graphClient
          .api(`/sites/${siteId}/lists`)
          .post(listDefinition);

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

      // Query alerts from each site
      for (const siteId of sitesToQuery) {
        try {
          const response = await this.graphClient
            .api(`/sites/${siteId}/lists/${this.alertsListName}/items`)
            .expand('fields,author')
            .get();

          const siteAlerts = response.value.map((item: any) => this.mapSharePointItemToAlert(item));
          allAlerts.push(...siteAlerts);
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
          AlertType: alert.AlertType,
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
        .expand('fields,author')
        .get();

      return this.mapSharePointItemToAlert(createdItem);
    } catch (error) {
      console.error('Failed to create alert:', error);
      throw error;
    }
  }

  /**
   * Update an existing alert
   */
  public async updateAlert(alertId: string, updates: Partial<IAlertItem>): Promise<IAlertItem> {
    try {
      const siteId = this.context.pageContext.site.id.toString();

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
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${alertId}/fields`)
        .patch(listItem.fields);

      // Get the updated item
      const updatedItem = await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${alertId}`)
        .expand('fields,author')
        .get();

      return this.mapSharePointItemToAlert(updatedItem);
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
      const siteId = this.context.pageContext.site.id.toString();

      await this.graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}/items/${alertId}`)
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
  private mapSharePointItemToAlert(item: any): IAlertItem {
    const fields = item.fields;
    
    // Create the original list item for multi-language support
    const originalListItem: IAlertListItem = {
      Id: parseInt(item.id.toString()),
      Title: fields.Title || '',
      Description: fields.Description || '',
      AlertType: fields.AlertType || '',
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
      id: item.id.toString(),
      title: fields.Title || '',
      description: fields.Description || '',
      AlertType: fields.AlertType || '',
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
      if (alert.scheduledStart && new Date(alert.scheduledStart) > now) {
        return false; // Not yet active
      }

      if (alert.scheduledEnd && new Date(alert.scheduledEnd) < now) {
        return false; // Already expired
      }

      return alert.status === 'Active' ||
        (alert.status === 'Scheduled' &&
          alert.scheduledStart &&
          new Date(alert.scheduledStart) <= now);
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