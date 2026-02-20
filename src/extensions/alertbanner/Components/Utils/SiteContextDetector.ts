import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { logger } from '../Services/LoggerService';

export interface ISiteContext {
  siteId: string;
  webId: string;
  siteUrl: string;
  siteName: string;
  siteType: 'regular' | 'hub' | 'homesite' | 'team' | 'communication';
  isHubSite: boolean;
  hubSiteId?: string;
  isHomesite: boolean;
  associatedSites: string[];
  tenantUrl: string;
  userPermissions: ISitePermissions;
  isRootSite: boolean;
}

export interface ISitePermissions {
  canCreateAlerts: boolean;
  canManageAlerts: boolean;
  canViewAlerts: boolean;
  permissionLevel: 'none' | 'read' | 'contribute' | 'design' | 'fullControl' | 'owner';
}

export interface ISiteOption {
  id: string;
  name: string;
  url: string;
  type: 'regular' | 'hub' | 'homesite' | 'team' | 'communication';
  isHub: boolean;
  isHomesite: boolean;
  lastModified: string;
  userPermissions: ISitePermissions;
  parentHubId?: string;
}

export interface ISiteValidationResult {
  siteId: string;
  siteName: string;
  hasAccess: boolean;
  canCreateAlerts: boolean;
  permissionLevel: string;
  error?: string;
}

export class SiteContextDetector {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private currentSiteContext: ISiteContext | null = null;

  constructor(graphClient: MSGraphClientV3, context: ApplicationCustomizerContext) {
    this.graphClient = graphClient;
    this.context = context;
  }

  public async getCurrentSiteContext(): Promise<ISiteContext> {
    if (this.currentSiteContext) {
      return this.currentSiteContext;
    }

    try {
      const siteId = this.context.pageContext.site.id.toString();
      const webId = this.context.pageContext.web.id.toString();
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const siteName = this.context.pageContext.web.title;
      const tenantUrl = `https://${new URL(siteUrl).hostname}`;

      const siteDetails = await this.graphClient
        .api(`/sites/${siteId}`)
        .expand('drive')
        .get();

      const hubInfo = await this.checkIfHubSite(siteId);

      const isHomesite = await this.checkIfHomesite(siteUrl, tenantUrl);

      const associatedSites = hubInfo.isHub ? await this.getAssociatedSites(siteId) : [];

      const siteType = this.determineSiteType(siteDetails, hubInfo.isHub, isHomesite);

      const userPermissions = await this.getUserPermissions(siteId);

      const isRootSite = this.isRootSiteCollection(siteUrl);

      this.currentSiteContext = {
        siteId,
        webId,
        siteUrl,
        siteName,
        siteType,
        isHubSite: hubInfo.isHub,
        hubSiteId: hubInfo.hubSiteId,
        isHomesite,
        associatedSites,
        tenantUrl,
        userPermissions,
        isRootSite
      };

      return this.currentSiteContext;
    } catch (error) {
      logger.error('SiteContextDetector', 'Failed to get site context', error);
      return {
        siteId: this.context.pageContext.site.id.toString(),
        webId: this.context.pageContext.web.id.toString(),
        siteUrl: this.context.pageContext.web.absoluteUrl,
        siteName: this.context.pageContext.web.title,
        siteType: 'regular',
        isHubSite: false,
        isHomesite: false,
        associatedSites: [],
        tenantUrl: `https://${new URL(this.context.pageContext.web.absoluteUrl).hostname}`,
        userPermissions: {
          canCreateAlerts: false,
          canManageAlerts: false,
          canViewAlerts: true,
          permissionLevel: 'read'
        },
        isRootSite: false
      };
    }
  }

  public async getAvailableSites(includePermissionCheck: boolean = true): Promise<ISiteOption[]> {
    try {
      const currentContext = await this.getCurrentSiteContext();
      const followedSites = await this.getFollowedSites();
      const hubSites = currentContext.isHubSite ?
        await this.getHubAssociatedSites(currentContext.siteId) : [];
      const recentSites = await this.getRecentSites();
      const allSites = new Map<string, ISiteOption>();

      [...followedSites, ...hubSites, ...recentSites].forEach(site => {
        if (!allSites.has(site.id)) {
          allSites.set(site.id, site);
        }
      });
      if (includePermissionCheck) {
        const sitesWithPermissions = await Promise.all(
          Array.from(allSites.values()).map(async site => {
            try {
              const permissions = await this.getUserPermissions(site.id);
              return {
                ...site,
                userPermissions: permissions
              };
            } catch (error) {
              logger.warn('SiteContextDetector', `Failed to check permissions for site ${site.id}`, error);
              return {
                ...site,
                userPermissions: {
                  canCreateAlerts: false,
                  canManageAlerts: false,
                  canViewAlerts: false,
                  permissionLevel: 'none' as const
                }
              };
            }
          })
        );
        return sitesWithPermissions;
      }

      return Array.from(allSites.values());
    } catch (error) {
      logger.error('SiteContextDetector', 'Failed to get available sites', error);
      return [];
    }
  }
  public async validateSiteAccess(siteIds: string[]): Promise<ISiteValidationResult[]> {
    const validationPromises = siteIds.map(async (siteId) => {
      try {
        const site = await this.graphClient
          .api(`/sites/${siteId}`)
          .select('id,displayName,webUrl')
          .get();

        const permissions = await this.getUserPermissions(siteId);

        return {
          siteId,
          siteName: site.displayName,
          hasAccess: permissions.canViewAlerts,
          canCreateAlerts: permissions.canCreateAlerts,
          permissionLevel: permissions.permissionLevel,
        };
      } catch (error) {
        return {
          siteId,
          siteName: 'Unknown Site',
          hasAccess: false,
          canCreateAlerts: false,
          permissionLevel: 'none',
          error: error.message
        };
      }
    });

    return Promise.all(validationPromises);
  }
  public async getSuggestedDistributionScopes(): Promise<{
    currentSite: ISiteOption;
    hubSites?: ISiteOption[];
    homesite?: ISiteOption;
    recentSites: ISiteOption[];
    followedSites: ISiteOption[];
  }> {
    const currentContext = await this.getCurrentSiteContext();
    const currentSite: ISiteOption = {
      id: currentContext.siteId,
      name: currentContext.siteName,
      url: currentContext.siteUrl,
      type: currentContext.siteType,
      isHub: currentContext.isHubSite,
      isHomesite: currentContext.isHomesite,
      lastModified: new Date().toISOString(),
      userPermissions: currentContext.userPermissions
    };

    const [recentSites, followedSites] = await Promise.all([
      this.getRecentSites(),
      this.getFollowedSites()
    ]);

    const enrichWithPermissions = async (sites: ISiteOption[]): Promise<ISiteOption[]> => {
      const enhancedSites = await Promise.all(sites.map(async site => {
        try {
          const permissions = await this.getUserPermissions(site.id);
          return { ...site, userPermissions: permissions };
        } catch (error) {
          logger.warn('SiteContextDetector', `Failed to evaluate permissions for site ${site.id}`, error);
          return {
            ...site,
            userPermissions: {
              canCreateAlerts: false,
              canManageAlerts: false,
              canViewAlerts: false,
              permissionLevel: 'none' as const
            }
          };
        }
      }));

      return enhancedSites.filter(site => site.userPermissions.canViewAlerts);
    };

    const [recentSitesWithPermissions, followedSitesWithPermissions] = await Promise.all([
      enrichWithPermissions(recentSites),
      enrichWithPermissions(followedSites)
    ]);

    const result: any = {
      currentSite,
      recentSites: recentSitesWithPermissions
        .filter(site => site.userPermissions.canCreateAlerts)
        .slice(0, 5),
      followedSites: followedSitesWithPermissions
        .filter(site => site.userPermissions.canCreateAlerts)
        .slice(0, 10)
    };

    if (currentContext.isHubSite) {
      result.hubSites = await this.getHubAssociatedSites(currentContext.siteId);
    }

    if (!currentContext.isHomesite) {
      const homesite = await this.getHomesite();
      if (homesite) {
        result.homesite = homesite;
      }
    }

    return result;
  }


  private async checkIfHubSite(siteId: string): Promise<{ isHub: boolean; hubSiteId?: string }> {
    try {
      const siteDetails = await this.graphClient
        .api(`/sites/${siteId}`)
        .select('sharepointIds')
        .get();

      const hasHubSiteId = siteDetails.sharepointIds?.hubSiteId;
      const isMarkedAsHub = Boolean(hasHubSiteId && hasHubSiteId === siteId);

      if (isMarkedAsHub) {
        return { isHub: true, hubSiteId: siteId };
      }

      if (hasHubSiteId) {
        return { isHub: false, hubSiteId: siteDetails.sharepointIds.hubSiteId };
      }

      if (this.context.pageContext.site.id.toString() === siteId) {
        const isHub = (this.context.pageContext.site as any).isHubSite === true;
        return { isHub };
      }

      return { isHub: false };
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not determine hub site status', error);
      return { isHub: false };
    }
  }

  private async checkIfHomesite(siteUrl: string, tenantUrl: string): Promise<boolean> {
    try {
      const url = new URL(siteUrl);
      const tenant = new URL(tenantUrl);

      const isRootSite = url.hostname === tenant.hostname &&
                        (url.pathname === '/' || url.pathname === '' || url.pathname === '/sites/root');

      if (!isRootSite) {
        return false;
      }

      try {
        const rootSite = await this.graphClient
          .api('/sites/root')
          .select('webUrl')
          .get();

        if (rootSite?.webUrl) {
          const homeSiteUrl = new URL(rootSite.webUrl);
          return url.hostname === homeSiteUrl.hostname && url.pathname === homeSiteUrl.pathname;
        }
      } catch (apiError) {
      }

      return isRootSite;
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not determine homesite status', error);
      return false;
    }
  }

  private async getAssociatedSites(hubSiteId: string): Promise<string[]> {
    try {
      const sites = await this.graphClient
        .api('/sites')
        .filter(`sharepointIds/hubSiteId eq '${hubSiteId}'`)
        .select('id,displayName,webUrl')
        .top(100)
        .get();

      if (sites?.value && Array.isArray(sites.value)) {
        return sites.value.map((site: any) => site.id);
      }

      return [];
    } catch (error) {
      if (!(error.statusCode === 400 || error.message?.includes('filter'))) {
        logger.warn('SiteContextDetector', 'Could not get associated sites', error);
      }
      return [];
    }
  }

  private determineSiteType(siteDetails: any, isHub: boolean, isHomesite: boolean): 'regular' | 'hub' | 'homesite' | 'team' | 'communication' {
    if (isHomesite) return 'homesite';
    if (isHub) return 'hub';

    if (siteDetails.webUrl?.includes('/teams/')) {
      return 'team';
    }

    if (siteDetails.description?.toLowerCase().includes('communication')) {
      return 'communication';
    }

    return 'regular';
  }

  private async getUserPermissions(siteId: string): Promise<ISitePermissions> {
    try {
      let hasWritePermission = false;
      let hasOwnerPermission = false;
      let permissionLevel: 'none' | 'read' | 'contribute' | 'design' | 'fullControl' | 'owner' = 'read';

      try {
        // First, try to get basic site information - if successful, user has at least read access
        await this.graphClient
          .api(`/sites/${siteId}`)
          .select('id,displayName')
          .get();

        try {
          await this.graphClient
            .api(`/sites/${siteId}/lists`)
            .select('id,displayName')
            .top(1)
            .get();

          // If we can read lists, user likely has contribute or higher permissions
          hasWritePermission = true;
          permissionLevel = 'contribute';

          // Additional check: try to access site columns (requires design or full control)
          try {
            await this.graphClient
              .api(`/sites/${siteId}/columns`)
              .select('id')
              .top(1)
              .get();

            hasOwnerPermission = true;
            permissionLevel = 'fullControl';
          } catch (columnError) {
          }
        } catch (listError) {
          permissionLevel = 'read';
        }

        return {
          canCreateAlerts: hasWritePermission,
          canManageAlerts: hasOwnerPermission,
          canViewAlerts: true,
          permissionLevel
        };
      } catch (siteError: any) {
        // Check if it's a 403 or 404 error - these are expected when user doesn't have access
        const statusCode = siteError?.statusCode || siteError?.status;
        if (statusCode === 403 || statusCode === 404) {
          return {
            canCreateAlerts: false,
            canManageAlerts: false,
            canViewAlerts: false,
            permissionLevel: 'none'
          };
        }

        // For other errors, log as warning
        logger.warn('SiteContextDetector', `Unexpected error checking permissions for site ${siteId}`, siteError);

        // For the current site, assume user has read permissions since they're viewing it
        const currentSiteId = this.context.pageContext.site.id.toString();
        if (siteId === currentSiteId) {
          return {
            canCreateAlerts: false,
            canManageAlerts: false,
            canViewAlerts: true,
            permissionLevel: 'read'
          };
        }

        return {
          canCreateAlerts: false,
          canManageAlerts: false,
          canViewAlerts: false,
          permissionLevel: 'none'
        };
      }
    } catch (error) {
      // This outer catch is for any unexpected errors in the function itself
      logger.warn('SiteContextDetector', `Failed to check permissions for site ${siteId}`, error);

      // For the current site, assume user has read permissions since they're viewing it
      const currentSiteId = this.context.pageContext.site.id.toString();
      if (siteId === currentSiteId) {
        return {
          canCreateAlerts: false,
          canManageAlerts: false,
          canViewAlerts: true,
          permissionLevel: 'read'
        };
      }

      return {
        canCreateAlerts: false,
        canManageAlerts: false,
        canViewAlerts: false,
        permissionLevel: 'none'
      };
    }
  }

  private isRootSiteCollection(siteUrl: string): boolean {
    try {
      const url = new URL(siteUrl);
      const path = url.pathname;
      return path === '/' || path === '' || path === '/sites/root';
    } catch {
      return false;
    }
  }

  private async getFollowedSites(): Promise<ISiteOption[]> {
    try {
      const followedSites = await this.graphClient
        .api('/me/followedSites')
        .select('id,displayName,webUrl,lastModifiedDateTime')
        .get();

      return followedSites.value.map((site: any) => ({
        id: site.id,
        name: site.displayName,
        url: site.webUrl,
        type: 'regular' as const,
        isHub: false,
        isHomesite: false,
        lastModified: site.lastModifiedDateTime,
        userPermissions: {
          canCreateAlerts: false, // Will be determined later if needed
          canManageAlerts: false,
          canViewAlerts: true,
          permissionLevel: 'read' as const
        }
      }));
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not get followed sites', error);
      return [];
    }
  }

  private async getHubAssociatedSites(hubSiteId: string): Promise<ISiteOption[]> {
    try {
      const siteIds = await this.getAssociatedSites(hubSiteId);

      if (siteIds.length === 0) {
        return [];
      }

      const siteDetailsPromises = siteIds.map(async (siteId) => {
        try {
          const site = await this.graphClient
            .api(`/sites/${siteId}`)
            .select('id,displayName,webUrl,lastModifiedDateTime')
            .get();

          return {
            id: site.id,
            name: site.displayName,
            url: site.webUrl,
            type: 'regular' as const,
            isHub: false,
            isHomesite: false,
            lastModified: site.lastModifiedDateTime,
            userPermissions: {
              canCreateAlerts: false,
              canManageAlerts: false,
              canViewAlerts: true,
              permissionLevel: 'read' as const
            },
            parentHubId: hubSiteId
          };
        } catch (error) {
          logger.warn('SiteContextDetector', `Could not get details for site ${siteId}`, error);
          return null;
        }
      });

      const siteDetails = await Promise.all(siteDetailsPromises);
      return siteDetails.filter((site) => site !== null) as ISiteOption[];
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not get hub associated sites', error);
      return [];
    }
  }

  private async getRecentSites(): Promise<ISiteOption[]> {
    try {
      // Get recent items - we'll filter out files ourselves since Graph API filter doesn't work reliably
      const recentSites = await this.graphClient
        .api('/me/insights/used')
        .top(20) // Get more items since we'll filter out files
        .get();

      const sitesWithUrls = recentSites.value.filter((item: any) => {
        const webUrl = item.resourceReference?.webUrl;
        if (!webUrl) return false;

        try {
          const pathname = new URL(webUrl).pathname;
          const isFile = /\.(pdf|doc|docx|xls|xlsx|ppt|pptx|txt|zip|stl|exe|jpg|jpeg|png|gif|bmp|svg|webp|sppkg|aspx|html|css|js|json|xml|csv|mp4|avi|mov|wmv|flv|wav|mp3|wma)$/i.test(pathname);
          return !isFile;
        } catch {
          return false;
        }
      });

      const sitesWithIds = await Promise.all(
        sitesWithUrls.slice(0, 10).map(async (item: any) => { // Limit to 10 after filtering
          const siteId = await this.extractSiteIdFromUrl(item.resourceReference.webUrl);
          return {
            id: siteId,
            name: item.resourceVisualization.title,
            url: item.resourceReference.webUrl,
            type: 'regular' as const,
            isHub: false,
            isHomesite: false,
            lastModified: item.lastUsed.lastAccessedDateTime,
            userPermissions: {
              canCreateAlerts: false,
              canManageAlerts: false,
              canViewAlerts: true,
              permissionLevel: 'read' as const
            }
          };
        })
      );

      return sitesWithIds.filter((site: ISiteOption) => site.id);
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not get recent sites', error);
      return [];
    }
  }

  private async getHomesite(): Promise<ISiteOption | null> {
    try {
      const homeSite = await this.graphClient
        .api('/sites/root')
        .select('id,displayName,webUrl')
        .get();

      return {
        id: homeSite.id,
        name: homeSite.displayName,
        url: homeSite.webUrl,
        type: 'homesite',
        isHub: false,
        isHomesite: true,
        lastModified: new Date().toISOString(),
        userPermissions: {
          canCreateAlerts: false,
          canManageAlerts: false,
          canViewAlerts: true,
          permissionLevel: 'read'
        }
      };
    } catch (error) {
      logger.warn('SiteContextDetector', 'Could not get homesite', error);
      return null;
    }
  }

  private async extractSiteIdFromUrl(url: string): Promise<string> {
    try {
      if (!url) return '';

      const pathname = new URL(url).pathname;
      const isFileUrl = /\.(pdf|doc|docx|xls|xlsx|ppt|pptx|txt|zip|stl|exe|jpg|jpeg|png|gif|bmp|svg|webp|sppkg|aspx|html|css|js|json|xml|csv|mp4|avi|mov|wmv|flv|wav|mp3|wma)$/i.test(pathname);

      if (isFileUrl) {
        return '';
      }

      const hostname = new URL(url).hostname;

      let apiPath = '';
      if (pathname === '/' || pathname === '') {
        apiPath = `/sites/${hostname}`;
      } else {
        apiPath = `/sites/${hostname}:${pathname}`;
      }

      const site = await this.graphClient
        .api(apiPath)
        .select('id')
        .get();

      return site.id || '';
    } catch (error) {
      logger.warn('SiteContextDetector', `Could not extract site ID from URL: ${url}`, error);
      return '';
    }
  }
}
