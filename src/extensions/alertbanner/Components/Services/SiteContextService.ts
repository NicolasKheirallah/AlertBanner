import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { logger } from './LoggerService';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { LIST_NAMES } from '../Utils/AppConstants';

export interface ISiteInfo {
  id: string;
  url: string;
  name: string;
  type: 'current' | 'hub' | 'home';
  hasAlertsList?: boolean;
  graphId?: string;
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
  private readonly alertsListName = LIST_NAMES.ALERTS;

  public static getInstance(
    context?: ApplicationCustomizerContext,
    graphClient?: MSGraphClientV3
  ): SiteContextService {
    if (!SiteContextService._instance) {
      SiteContextService._instance = new SiteContextService(context, graphClient);
    } else {
      // Update stale dependencies if provided
      if (context) {
        SiteContextService._instance._context = context;
      }
      if (graphClient) {
        SiteContextService._instance._graphClient = graphClient;
      }
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
      const hostname = new URL(this._context.pageContext.web.absoluteUrl).hostname;
      const siteId = this._context.pageContext.site.id.toString();
      const siteCollectionUrl = this._context.pageContext.site.absoluteUrl || this._context.pageContext.web.absoluteUrl;
      const graphId = this.buildGraphSiteIdentifier(siteCollectionUrl);

      // Get current site info
      this._currentSiteInfo = {
        id: siteId,
        url: siteCollectionUrl,
        name: this._context.pageContext.web.title,
        type: 'current',
        graphId: graphId
      };

      // Detect home site
      await this.detectHomeSite();

      // Detect hub site if current site is connected to a hub
      await this.detectHubSite();

      // Check alert lists for all sites
      await this.checkAlertLists();

    } catch (error) {
      logger.error('SiteContextService', 'Failed to detect site context', error);
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
              type: 'home',
              graphId: homeSiteResponse.id
            };
          }
        } catch (rootError) {
          // Root site not accessible or not home site
        }
      }

      // If still no home site found, try alternative search method
      if (!this._homeSiteInfo) {
        await this.searchForHomeSite();
      }
    } catch (error) {
      logger.warn('SiteContextService', 'Could not detect home site through organization API', error);
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
            query: {
              queryString: 'IsHomeSite:true OR SiteTemplate:SITEPAGEPUBLISHING'
            },
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
              type: 'home',
              graphId: site.id || site.siteId
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
            type: 'home',
            graphId: firstSite.id || firstSite.siteId
          };
        }
      }
    } catch (error) {
      logger.warn('SiteContextService', 'Could not find home site through search API', error);
      
      // Final fallback: try to find sites the user can access and look for patterns
      try {
        const sitesResponse = await this._graphClient
          .api('/sites?search=*')
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
              type: 'home',
              graphId: potentialHomeSite.id
            };
          }
        }
      } catch (sitesError) {
        logger.warn('SiteContextService', 'Could not access sites collection', sitesError);
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
          type: 'hub',
          graphId: hubResponse.id
        };
      }
    } catch (error) {
      logger.warn('SiteContextService', 'Could not detect hub site', error);
    }
  }

  /**
   * Check if alert lists exist on all relevant sites
   */
  private async checkAlertLists(): Promise<void> {
    const sites = [this._currentSiteInfo, this._hubSiteInfo, this._homeSiteInfo].filter(Boolean) as ISiteInfo[];

    for (const site of sites) {
      if (site) {
        try {
          site.hasAlertsList = await this.checkAlertListExists(site);
        } catch (error) {
          logger.warn('SiteContextService', `Failed to check alerts list for site ${site.name}`, error);
          site.hasAlertsList = false;
        }
      }
    }
  }

  /**
   * Check if alerts list exists on a specific site
   */
  public async checkAlertListExists(site: ISiteInfo): Promise<boolean> {
    try {
      const graphSiteId = site.graphId ?? this.buildGraphSiteIdentifier(site.url);
      const response = await this._graphClient
        .api(`/sites/${graphSiteId}/lists`)
        .filter(`displayName eq '${this.alertsListName}'`)
        .select('id')
        .top(1)
        .get();

      return Array.isArray(response?.value) && response.value.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * Get detailed status of alerts list on a site
   */
  public async getAlertListStatus(site: ISiteInfo): Promise<IAlertListStatus> {
    try {
      const graphSiteId = site.graphId ?? this.buildGraphSiteIdentifier(site.url);
      // Try to access the list
      await this._graphClient
        .api(`/sites/${graphSiteId}/lists`)
        .filter(`displayName eq '${this.alertsListName}'`)
        .top(1)
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
            .api(`/sites/${site.graphId ?? this.buildGraphSiteIdentifier(site.url)}/lists`)
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
   * Create alerts list on a specific site using SharePointAlertService
   */
  public async createAlertsList(siteId: string, selectedLanguages?: string[]): Promise<boolean> {
    try {
      // Import SharePointAlertService dynamically to avoid circular dependency
      const { SharePointAlertService } = await import('./SharePointAlertService');
      const alertService = new SharePointAlertService(this._graphClient, this._context);
      
      // Pass siteId directly to initializeLists and addLanguageSupport
      await alertService.initializeLists(siteId);
      
      if (selectedLanguages && selectedLanguages.length > 0) {
        await alertService.updateSupportedLanguages(siteId, selectedLanguages);
      }
      
      // Update the hasAlertsList flag for the site
      const sites = [this._currentSiteInfo, this._hubSiteInfo, this._homeSiteInfo];
      const targetSite = sites.find(s => s && s.id === siteId);
      if (targetSite) {
        targetSite.hasAlertsList = true;
      }
      
      return true;
    } catch (error) {
      logger.error('SiteContextService', `Failed to create alerts list on site ${siteId}`, error);
      
      // Provide more detailed error messages
      if (error.message?.includes('PERMISSION_DENIED')) {
        throw new Error(`PERMISSION_DENIED: Cannot create alerts list on site ${siteId}. User lacks required permissions.`);
      } else if (error.message?.includes('CRITICAL_COLUMNS_FAILED')) {
        throw new Error(`LIST_INCOMPLETE: Alerts list created but some critical columns failed. ${error.message}`);
      } else {
        throw new Error(`LIST_CREATION_FAILED: ${error.message || 'Unknown error during list creation'}`);
      }
    }
  }

  /**
   * Get supported languages for a specific site's alerts list
   */
  public async getSupportedLanguagesForSite(siteId: string): Promise<string[]> {
    try {
      // Import SharePointAlertService dynamically to avoid circular dependency
      const { SharePointAlertService } = await import('./SharePointAlertService');
      const alertService = new SharePointAlertService(this._graphClient, this._context);
      
      // Pass siteId directly to getSupportedLanguages
      return await alertService.getSupportedLanguages(siteId);
    } catch (error) {
      logger.warn('SiteContextService', `Failed to get supported languages for site ${siteId}`, error);
      return ['en-us']; // Default fallback
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
  /**
   * Get sites that should show alerts for current user context
   */
  public getAlertSourceSites(): string[] {
    const siteMap = new Map<string, string>(); // Map: GUID -> siteId (for deduplication)

    // Helper to add site to map
    const addSite = (siteInfo: ISiteInfo | null) => {
      if (siteInfo && siteInfo.hasAlertsList) {
        const guid = siteInfo.id;
        const graphId = siteInfo.graphId ?? guid ?? this.buildGraphSiteIdentifier(siteInfo.url);
        if (guid && !siteMap.has(guid)) {
          siteMap.set(guid, graphId);
        }
      }
    };

    // 1. Current Site
    addSite(this._currentSiteInfo);

    // 2. Hub Site (if connected and different from Current)
    if (this._hubSiteInfo?.id !== this._currentSiteInfo?.id) {
      addSite(this._hubSiteInfo);
    }

    // 3. Home Site (Always, if different from Current and Hub)
    if (this._homeSiteInfo?.id !== this._currentSiteInfo?.id && 
        this._homeSiteInfo?.id !== this._hubSiteInfo?.id) {
      addSite(this._homeSiteInfo);
    }

    // Return unique site IDs
    return Array.from(siteMap.values());
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
   * Get application context
   */
  public getContext(): ApplicationCustomizerContext {
    return this._context;
  }

  /**
   * Get Microsoft Graph client
   */
  public async getGraphClient(): Promise<MSGraphClientV3> {
    return this._graphClient;
  }

  private buildGraphSiteIdentifier(siteUrl: string): string {
    const normalizedUrl = new URL(siteUrl);
    let path = normalizedUrl.pathname || '';

    // Normalize redundant slashes and remove trailing slash except for root
    path = path.replace(/\/+/g, '/');
    if (path !== '/' && path.endsWith('/')) {
      path = path.slice(0, -1);
    }

    if (!path || path === '/') {
      return normalizedUrl.hostname;
    }

    return `${normalizedUrl.hostname}:${path}`;
  }

  /**
   * Refresh site context (useful after list creation)
   */
  public async refresh(): Promise<void> {
    await this.detectSiteContext();
  }
}
