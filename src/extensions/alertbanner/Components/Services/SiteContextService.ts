import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export interface ISiteInfo {
  id: string;
  url: string;
  name: string;
  type: 'current' | 'hub' | 'home';
  hasAlertsList?: boolean;
}

export interface IAlertListStatus {
  exists: boolean;
  canAccess: boolean;
  canCreate: boolean;
  error?: string;
}

export class SiteContextService {
  private static _instance: SiteContextService;
  private _context: ApplicationCustomizerContext;
  private _graphClient: MSGraphClientV3;
  private _homeSiteInfo: ISiteInfo | null = null;
  private _hubSiteInfo: ISiteInfo | null = null;
  private _currentSiteInfo: ISiteInfo | null = null;
  private readonly alertsListName = 'Alerts';

  public static getInstance(
    context?: ApplicationCustomizerContext,
    graphClient?: MSGraphClientV3
  ): SiteContextService {
    if (!SiteContextService._instance) {
      SiteContextService._instance = new SiteContextService(context, graphClient);
    }
    return SiteContextService._instance;
  }

  private constructor(
    context?: ApplicationCustomizerContext,
    graphClient?: MSGraphClientV3
  ) {
    if (context) this._context = context;
    if (graphClient) this._graphClient = graphClient;
  }

  /**
   * Initialize the service and detect site context
   */
  public async initialize(): Promise<void> {
    if (!this._context || !this._graphClient) {
      throw new Error('SiteContextService requires context and graphClient');
    }

    await this.detectSiteContext();
  }

  /**
   * Detect current site, hub site, and home site
   */
  private async detectSiteContext(): Promise<void> {
    try {
      // Get current site info
      this._currentSiteInfo = {
        id: this._context.pageContext.site.id.toString(),
        url: this._context.pageContext.site.absoluteUrl,
        name: (this._context.pageContext.site as any).displayName || 'Current Site',
        type: 'current'
      };

      // Detect home site
      await this.detectHomeSite();

      // Detect hub site if current site is connected to a hub
      await this.detectHubSite();

      // Check alert lists for all sites
      await this.checkAlertLists();

    } catch (error) {
      console.error('Failed to detect site context:', error);
    }
  }

  /**
   * Detect the tenant's home site using user-accessible APIs
   */
  private async detectHomeSite(): Promise<void> {
    try {
      // Use the more accessible approach: try to get organization settings
      // This doesn't require SharePoint Admin permissions
      const orgResponse = await this._graphClient
        .api('/organization')
        .get();

      // Try to get tenant information which might include home site
      if (orgResponse?.value?.[0]) {
        const tenantId = orgResponse.value[0].id;
        
        // Try to find home site through organization information
        try {
          const homeSiteResponse = await this._graphClient
            .api('/sites/root')
            .get();

          // Check if root site is configured as home site
          if (homeSiteResponse?.sharepointIds?.tenantId === tenantId) {
            this._homeSiteInfo = {
              id: homeSiteResponse.id,
              url: homeSiteResponse.webUrl,
              name: (homeSiteResponse as any).displayName || (homeSiteResponse as any).name || 'Home Site',
              type: 'home'
            };
          }
        } catch (rootError) {
          console.log('Root site not accessible or not home site');
        }
      }

      // If still no home site found, try alternative search method
      if (!this._homeSiteInfo) {
        await this.searchForHomeSite();
      }
    } catch (error) {
      console.warn('Could not detect home site through organization API:', error);
      // Try alternative method using search
      await this.searchForHomeSite();
    }
  }

  /**
   * Alternative method to find home site using user-accessible APIs
   */
  private async searchForHomeSite(): Promise<void> {
    try {
      // Try to use Microsoft Search API which is more accessible
      const searchResponse = await this._graphClient
        .api('/search/query')
        .post({
          requests: [{
            entityTypes: ['site'],
            query: 'IsHomeSite:true OR SiteTemplate:SITEPAGEPUBLISHING',
            from: 0,
            size: 5
          }]
        });

      const results = searchResponse.value[0]?.hitsContainers[0]?.hits;
      if (results && results.length > 0) {
        // Look for a site that might be the home site
        for (const result of results) {
          const site = result.resource;
          // Check if this looks like a home site (typically has specific characteristics)
          if (site.webUrl && (site.webUrl.includes('/sites/home') || 
                             site.webUrl.includes('/sites/intranet') ||
                             site.displayName?.toLowerCase().includes('home') ||
                             site.displayName?.toLowerCase().includes('intranet'))) {
            this._homeSiteInfo = {
              id: site.id || site.siteId,
              url: site.webUrl,
              name: site.displayName || site.name || 'Home Site',
              type: 'home'
            };
            break;
          }
        }

        // If no obvious home site found, use the first result as a fallback
        if (!this._homeSiteInfo && results.length > 0) {
          const firstSite = results[0].resource;
          this._homeSiteInfo = {
            id: firstSite.id || firstSite.siteId,
            url: firstSite.webUrl,
            name: firstSite.displayName || firstSite.name || 'Tenant Root Site',
            type: 'home'
          };
        }
      }
    } catch (error) {
      console.warn('Could not find home site through search API:', error);
      
      // Final fallback: try to find sites the user can access and look for patterns
      try {
        const sitesResponse = await this._graphClient
          .api('/sites')
          .filter("siteCollection/root ne null")
          .top(10)
          .get();

        if (sitesResponse?.value?.length > 0) {
          // Look for a site that might be home site based on naming patterns
          const potentialHomeSite = sitesResponse.value.find((site: any) => 
            site.webUrl?.includes('/sites/home') || 
            site.webUrl?.includes('/sites/intranet') ||
            site.displayName?.toLowerCase().includes('home') ||
            site.displayName?.toLowerCase().includes('intranet')
          );

          if (potentialHomeSite) {
            this._homeSiteInfo = {
              id: potentialHomeSite.id,
              url: potentialHomeSite.webUrl,
              name: potentialHomeSite.displayName || 'Home Site',
              type: 'home'
            };
          }
        }
      } catch (sitesError) {
        console.warn('Could not access sites collection:', sitesError);
        // At this point, we'll proceed without home site detection
      }
    }
  }

  /**
   * Detect hub site if current site is connected to one
   */
  private async detectHubSite(): Promise<void> {
    try {
      if (this._context.pageContext.legacyPageContext.hubSiteId) {
        // Current site is connected to a hub
        const hubSiteId = this._context.pageContext.legacyPageContext.hubSiteId;
        const hubResponse = await this._graphClient
          .api(`/sites/${hubSiteId}`)
          .get();

        this._hubSiteInfo = {
          id: hubSiteId,
          url: hubResponse.webUrl,
          name: (hubResponse as any).displayName || (hubResponse as any).name || 'Hub Site',
          type: 'hub'
        };
      }
    } catch (error) {
      console.warn('Could not detect hub site:', error);
    }
  }

  /**
   * Check if alert lists exist on all relevant sites
   */
  private async checkAlertLists(): Promise<void> {
    const sites = [this._currentSiteInfo, this._hubSiteInfo, this._homeSiteInfo].filter(Boolean);
    
    for (const site of sites) {
      if (site) {
        try {
          site.hasAlertsList = await this.checkAlertListExists(site.id);
        } catch (error) {
          console.warn(`Failed to check alerts list for site ${site.name}:`, error);
          site.hasAlertsList = false;
        }
      }
    }
  }

  /**
   * Check if alerts list exists on a specific site
   */
  public async checkAlertListExists(siteId: string): Promise<boolean> {
    try {
      await this._graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}`)
        .get();
      return true;
    } catch (error) {
      return false;
    }
  }

  /**
   * Get detailed status of alerts list on a site
   */
  public async getAlertListStatus(siteId: string): Promise<IAlertListStatus> {
    try {
      // Try to access the list
      await this._graphClient
        .api(`/sites/${siteId}/lists/${this.alertsListName}`)
        .get();
      
      return {
        exists: true,
        canAccess: true,
        canCreate: false // Already exists
      };
    } catch (error) {
      if (error.message?.includes('404') || error.message?.includes('not found')) {
        // List doesn't exist, check if we can create it
        try {
          // Test permissions by trying to get all lists
          await this._graphClient
            .api(`/sites/${siteId}/lists`)
            .select('id')
            .top(1)
            .get();

          return {
            exists: false,
            canAccess: true,
            canCreate: true
          };
        } catch (permError) {
          return {
            exists: false,
            canAccess: false,
            canCreate: false,
            error: 'Insufficient permissions to access or create lists'
          };
        }
      } else if (error.message?.includes('403') || error.message?.includes('Access denied')) {
        return {
          exists: true, // Assume it exists but we can't access it
          canAccess: false,
          canCreate: false,
          error: 'Access denied to alerts list'
        };
      }

      return {
        exists: false,
        canAccess: false,
        canCreate: false,
        error: error.message
      };
    }
  }

  /**
   * Create alerts list on a specific site
   */
  public async createAlertsList(siteId: string): Promise<boolean> {
    try {
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
          // Note: Multi-language fields will be added dynamically based on user selection
        ]
      };

      await this._graphClient
        .api(`/sites/${siteId}/lists`)
        .post(listDefinition);

      console.log(`Successfully created alerts list on site ${siteId}`);
      return true;
    } catch (error) {
      console.error(`Failed to create alerts list on site ${siteId}:`, error);
      return false;
    }
  }

  /**
   * Get all relevant sites in hierarchical order
   */
  public getSitesHierarchy(): ISiteInfo[] {
    const sites: ISiteInfo[] = [];
    
    // Add in priority order: Home → Hub → Current
    if (this._homeSiteInfo) sites.push(this._homeSiteInfo);
    if (this._hubSiteInfo && this._hubSiteInfo.id !== this._homeSiteInfo?.id) {
      sites.push(this._hubSiteInfo);
    }
    if (this._currentSiteInfo && 
        this._currentSiteInfo.id !== this._homeSiteInfo?.id && 
        this._currentSiteInfo.id !== this._hubSiteInfo?.id) {
      sites.push(this._currentSiteInfo);
    }

    return sites;
  }

  /**
   * Get sites that should show alerts for current user context
   */
  public getAlertSourceSites(): string[] {
    const siteIds: string[] = [];
    
    // Always include home site alerts (shown everywhere)
    if (this._homeSiteInfo?.hasAlertsList) {
      siteIds.push(this._homeSiteInfo.id);
    }

    // Include hub site alerts if current site is connected to hub
    if (this._hubSiteInfo?.hasAlertsList && this._context.pageContext.legacyPageContext.hubSiteId) {
      siteIds.push(this._hubSiteInfo.id);
    }

    // Always include current site alerts
    if (this._currentSiteInfo?.hasAlertsList) {
      siteIds.push(this._currentSiteInfo.id);
    }

    return siteIds;
  }

  /**
   * Get current site info
   */
  public getCurrentSite(): ISiteInfo | null {
    return this._currentSiteInfo;
  }

  /**
   * Get hub site info
   */
  public getHubSite(): ISiteInfo | null {
    return this._hubSiteInfo;
  }

  /**
   * Get home site info
   */
  public getHomeSite(): ISiteInfo | null {
    return this._homeSiteInfo;
  }

  /**
   * Utility methods
   */
  // Removed unused extractHostnameFromUrl and extractPathFromUrl methods

  /**
   * Refresh site context (useful after list creation)
   */
  public async refresh(): Promise<void> {
    await this.detectSiteContext();
  }
}