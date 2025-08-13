import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { AlertPriority, NotificationType, IAlertType } from "../Alerts/IAlerts";

export interface IAlertItem {
  id: string;
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl?: string;
  linkDescription?: string;
  targetSites: string[];
  status: 'Active' | 'Expired' | 'Scheduled';
  createdDate: string;
  createdBy: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  metadata?: any;
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
}

export class SharePointAlertService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private alertsListName = 'AlertBannerAlerts';
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
      
      // Check if alerts list exists
      const alertsListCreated = await this.ensureAlertsList(siteId);
      const typesListCreated = await this.ensureAlertTypesList(siteId);
      
      // Lists created successfully
      if (alertsListCreated) {
        console.log('Alert Banner alerts list created successfully');
      }
      
      if (typesListCreated) {
        console.log('Alert Banner types list created successfully');
      }
    } catch (error) {
      // Enhanced error handling for common permission issues
      if (error.message?.includes('Access denied') || error.message?.includes('403')) {
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
              displayAs: 'dropDownMenu'
            }
          },
          {
            name: 'IsPinned',
            boolean: {}
          },
          {
            name: 'NotificationType',
            choice: {
              choices: ['None', 'Browser', 'Email', 'Both'],
              displayAs: 'dropDownMenu'
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
              displayAs: 'dropDownMenu'
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
        ]
      };

      await this.graphClient
        .api(`/sites/${siteId}/lists`)
        .post(listDefinition);
      
      return true; // List was created
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
            number: { decimalPlaces: 0 }
          }
        ]
      };

      await this.graphClient
        .api(`/sites/${siteId}/lists`)
        .post(listDefinition);
      
      return true; // List was created
    }
  }

  /**
   * Get all alerts from SharePoint
   */
  public async getAlerts(siteIds?: string[]): Promise<IAlertItem[]> {
    try {
      const currentSiteId = this.context.pageContext.site.id.toString();
      const sitesToQuery = siteIds || [currentSiteId];
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
      metadata: fields.Metadata ? JSON.parse(fields.Metadata) : undefined
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
}