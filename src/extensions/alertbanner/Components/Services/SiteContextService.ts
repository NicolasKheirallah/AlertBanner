import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { logger } from './LoggerService';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { LIST_NAMES } from '../Utils/AppConstants';
import { SiteIdUtils } from "../Utils/SiteIdUtils";

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

interface ICachedSiteContext {
  timestamp: number;
  currentSiteInfo: ISiteInfo | null;
  hubSiteInfo: ISiteInfo | null;
  homeSiteInfo: ISiteInfo | null;
}

export class SiteContextService {
  private static _instance: SiteContextService;
  private _context: ApplicationCustomizerContext;
  private _graphClient: MSGraphClientV3;
  private _homeSiteInfo: ISiteInfo | null = null;
  private _hubSiteInfo: ISiteInfo | null = null;
  private _currentSiteInfo: ISiteInfo | null = null;
  private readonly alertsListName = LIST_NAMES.ALERTS;
  private _isInitializing: boolean = false;
  private _initPromise: Promise<void> | null = null;
  
  private static readonly CACHE_KEY = 'AlertBanner_SiteContext';
  private static readonly CACHE_DURATION_MS = 5 * 60 * 1000; // 5 minutes

  public static getInstance(
    context?: ApplicationCustomizerContext,
    graphClient?: MSGraphClientV3
  ): SiteContextService {
    if (!SiteContextService._instance) {
      SiteContextService._instance = new SiteContextService(context, graphClient);
    } else {
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

  public async initialize(): Promise<void> {
    if (!this._context || !this._graphClient) {
      throw new Error('SiteContextService requires context and graphClient');
    }

    if (this._isInitializing && this._initPromise) {
      return this._initPromise;
    }

    const cached = this._loadFromCache();
    if (cached && !this._isCacheExpired(cached)) {
      logger.debug('SiteContextService', 'Using cached site context');
      this._currentSiteInfo = cached.currentSiteInfo;
      this._hubSiteInfo = cached.hubSiteInfo;
      this._homeSiteInfo = cached.homeSiteInfo;
      return;
    }

    this._isInitializing = true;
    this._initPromise = this._doInitialization();
    
    try {
      await this._initPromise;
    } finally {
      this._isInitializing = false;
      this._initPromise = null;
    }
  }

  private async _doInitialization(): Promise<void> {
    const initStartTime = performance.now();
    
    try {
      this._initializeCurrentSite();

      const [homeSiteResult, hubSiteResult] = await Promise.all([
        this._detectHomeSiteWithTimeout(),
        this._detectHubSiteWithTimeout(),
      ]);

      this._homeSiteInfo = homeSiteResult;
      this._hubSiteInfo = hubSiteResult;

      await this._checkAlertListsParallel();

      this._saveToCache();

      logger.info('SiteContextService', 'Initialization complete', {
        durationMs: Math.round(performance.now() - initStartTime),
        hasHomeSite: !!this._homeSiteInfo,
        hasHubSite: !!this._hubSiteInfo,
      });
    } catch (error) {
      logger.error('SiteContextService', 'Failed to detect site context', error);
      if (!this._currentSiteInfo) {
        this._initializeCurrentSite();
      }
    }
  }

  private _initializeCurrentSite(): void {
    const siteCollectionUrl = this._context.pageContext.site.absoluteUrl || this._context.pageContext.web.absoluteUrl;
    this._currentSiteInfo = {
      id: this._context.pageContext.site.id.toString(),
      url: siteCollectionUrl,
      name: this._context.pageContext.web.title,
      type: 'current',
      graphId: this._buildGraphSiteIdentifier(siteCollectionUrl)
    };
  }

  private async _detectHomeSiteWithTimeout(): Promise<ISiteInfo | null> {
    return Promise.race([
      this._detectHomeSiteFast(),
      new Promise<null>((resolve) => 
        setTimeout(() => {
          logger.warn('SiteContextService', 'Home site detection timed out');
          resolve(null);
        }, 1500)
      ),
    ]);
  }

  private async _detectHomeSiteFast(): Promise<ISiteInfo | null> {
    try {
      const homeSiteResponse = await this._graphClient
        .api('/sites/root')
        .get();

      if (homeSiteResponse?.webUrl) {
        return {
          id: homeSiteResponse.id,
          url: homeSiteResponse.webUrl,
          name: (homeSiteResponse as any).displayName || (homeSiteResponse as any).name || 'Home Site',
          type: 'home',
          graphId: homeSiteResponse.id
        };
      }
    } catch (rootError) {
    }

    try {
      const searchResponse = await this._graphClient
        .api('/search/query')
        .post({
          requests: [{
            entityTypes: ['site'],
            query: { queryString: 'IsHomeSite:true' },
            from: 0,
            size: 1
          }]
        });

      const result = searchResponse.value[0]?.hitsContainers[0]?.hits?.[0];
      if (result?.resource) {
        const site = result.resource;
        return {
          id: site.id || site.siteId,
          url: site.webUrl,
          name: site.displayName || site.name || 'Home Site',
          type: 'home',
          graphId: site.id || site.siteId
        };
      }
    } catch (searchError) {
    }

    return null;
  }

  private async _detectHubSiteWithTimeout(): Promise<ISiteInfo | null> {
    return Promise.race([
      this._detectHubSite(),
      new Promise<null>((resolve) => 
        setTimeout(() => {
          logger.warn('SiteContextService', 'Hub site detection timed out');
          resolve(null);
        }, 1000)
      ),
    ]);
  }

  private async _detectHubSite(): Promise<ISiteInfo | null> {
    try {
      const hubSiteId = this._context.pageContext.legacyPageContext.hubSiteId;
      if (!hubSiteId) {
        return null;
      }

      const hubResponse = await this._graphClient
        .api(`/sites/${hubSiteId}`)
        .get();

      return {
        id: hubSiteId,
        url: hubResponse.webUrl,
        name: (hubResponse as any).displayName || (hubResponse as any).name || 'Hub Site',
        type: 'hub',
        graphId: hubResponse.id
      };
    } catch (error) {
      logger.warn('SiteContextService', 'Could not detect hub site', error);
      return null;
    }
  }

  private async _checkAlertListsParallel(): Promise<void> {
    const sites = [this._currentSiteInfo, this._hubSiteInfo, this._homeSiteInfo]
      .filter((s): s is ISiteInfo => s !== null);

    await Promise.all(
      sites.map(async (site) => {
        try {
          site.hasAlertsList = await this._checkAlertListExists(site);
        } catch (error) {
          logger.warn('SiteContextService', `Failed to check alerts list for ${site.name}`, error);
          site.hasAlertsList = false;
        }
      })
    );
  }

  private async _checkAlertListExists(site: ISiteInfo): Promise<boolean> {
    try {
      const graphSiteId = site.graphId ?? this._buildGraphSiteIdentifier(site.url);
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

  private _loadFromCache(): ICachedSiteContext | null {
    try {
      const cached = localStorage.getItem(SiteContextService.CACHE_KEY);
      if (cached) {
        return JSON.parse(cached) as ICachedSiteContext;
      }
    } catch {
    }
    return null;
  }

  private _saveToCache(): void {
    try {
      const cacheData: ICachedSiteContext = {
        timestamp: Date.now(),
        currentSiteInfo: this._currentSiteInfo,
        hubSiteInfo: this._hubSiteInfo,
        homeSiteInfo: this._homeSiteInfo,
      };
      localStorage.setItem(SiteContextService.CACHE_KEY, JSON.stringify(cacheData));
    } catch {
    }
  }

  private _isCacheExpired(cached: ICachedSiteContext): boolean {
    return Date.now() - cached.timestamp > SiteContextService.CACHE_DURATION_MS;
  }

  public async checkAlertListExists(site: ISiteInfo): Promise<boolean> {
    return this._checkAlertListExists(site);
  }

  public async getAlertListStatus(site: ISiteInfo): Promise<IAlertListStatus> {
    try {
      const graphSiteId = site.graphId ?? this._buildGraphSiteIdentifier(site.url);
      await this._graphClient
        .api(`/sites/${graphSiteId}/lists`)
        .filter(`displayName eq '${this.alertsListName}'`)
        .top(1)
        .get();

      return { exists: true, canAccess: true, canCreate: false };
    } catch (error: any) {
      if (error.message?.includes('404') || error.message?.includes('not found')) {
        try {
          await this._graphClient
            .api(`/sites/${site.graphId ?? this._buildGraphSiteIdentifier(site.url)}/lists`)
            .select('id')
            .top(1)
            .get();

          return { exists: false, canAccess: true, canCreate: true };
        } catch (permError) {
          return { exists: false, canAccess: false, canCreate: false, error: 'Insufficient permissions' };
        }
      } else if (error.message?.includes('403') || error.message?.includes('Access denied')) {
        return { exists: true, canAccess: false, canCreate: false, error: 'Access denied' };
      }

      return { exists: false, canAccess: false, canCreate: false, error: error.message };
    }
  }

  public async createAlertsList(siteId: string, selectedLanguages?: string[]): Promise<boolean> {
    try {
      const { SharePointAlertService } = await import('./SharePointAlertService');
      const alertService = new SharePointAlertService(this._graphClient, this._context);
      
      await alertService.initializeLists(siteId);
      
      if (selectedLanguages && selectedLanguages.length > 0) {
        await alertService.updateSupportedLanguages(siteId, selectedLanguages);
      }
      
      const sites = [this._currentSiteInfo, this._hubSiteInfo, this._homeSiteInfo];
      const targetSite = sites.find(s => s && s.id === siteId);
      if (targetSite) {
        targetSite.hasAlertsList = true;
      }
      
      this._clearCache();
      
      return true;
    } catch (error: any) {
      logger.error('SiteContextService', `Failed to create alerts list on site ${siteId}`, error);
      
      if (error.message?.includes('PERMISSION_DENIED')) {
        throw new Error(`PERMISSION_DENIED: Cannot create alerts list on site ${siteId}. User lacks required permissions.`);
      } else if (error.message?.includes('CRITICAL_COLUMNS_FAILED')) {
        throw new Error(`LIST_INCOMPLETE: Alerts list created but some critical columns failed. ${error.message}`);
      } else {
        throw new Error(`LIST_CREATION_FAILED: ${error.message || 'Unknown error during list creation'}`);
      }
    }
  }

  public async getSupportedLanguagesForSite(siteId: string): Promise<string[]> {
    try {
      const { SharePointAlertService } = await import('./SharePointAlertService');
      const alertService = new SharePointAlertService(this._graphClient, this._context);
      return await alertService.getSupportedLanguages(siteId);
    } catch (error) {
      logger.warn('SiteContextService', `Failed to get supported languages for site ${siteId}`, error);
      return ['en-us'];
    }
  }

  public getSitesHierarchy(): ISiteInfo[] {
    const sites: ISiteInfo[] = [];
    
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

  public getAlertSourceSites(): string[] {
    const siteMap = new Map<string, string>();

    const addSite = (siteInfo: ISiteInfo | null) => {
      if (siteInfo && siteInfo.hasAlertsList) {
        const idGuid = SiteIdUtils.extractGuidFromGraphId(siteInfo.id) ||
          (SiteIdUtils.isGuid(siteInfo.id) ? SiteIdUtils.normalizeGuid(siteInfo.id) : "");
        const graphGuid = SiteIdUtils.extractGuidFromGraphId(siteInfo.graphId || "") ||
          (SiteIdUtils.isGuid(siteInfo.graphId || "") ? SiteIdUtils.normalizeGuid(siteInfo.graphId || "") : "");

        const dedupKey = idGuid || graphGuid ||
          (siteInfo.url || "").toLowerCase().replace(/\/$/, "") ||
          (siteInfo.id || "").toLowerCase();

        if (!dedupKey || siteMap.has(dedupKey)) return;

        const canonicalIdentifier = idGuid || graphGuid || siteInfo.graphId ||
          siteInfo.id || this._buildGraphSiteIdentifier(siteInfo.url);

        siteMap.set(dedupKey, canonicalIdentifier);
      }
    };

    addSite(this._currentSiteInfo);
    if (this._hubSiteInfo?.id !== this._currentSiteInfo?.id) addSite(this._hubSiteInfo);
    if (this._homeSiteInfo?.id !== this._currentSiteInfo?.id && 
        this._homeSiteInfo?.id !== this._hubSiteInfo?.id) {
      addSite(this._homeSiteInfo);
    }

    return Array.from(siteMap.values());
  }

  public getCurrentSite(): ISiteInfo | null {
    return this._currentSiteInfo;
  }

  public getHubSite(): ISiteInfo | null {
    return this._hubSiteInfo;
  }
 
  public getHomeSite(): ISiteInfo | null {
    return this._homeSiteInfo;
  }

  public getContext(): ApplicationCustomizerContext {
    return this._context;
  }

  public async getGraphClient(): Promise<MSGraphClientV3> {
    return this._graphClient;
  }

  private _buildGraphSiteIdentifier(siteUrl: string): string {
    const normalizedUrl = new URL(siteUrl);
    let path = normalizedUrl.pathname || '';
    path = path.replace(/\/+/g, '/');
    if (path !== '/' && path.endsWith('/')) {
      path = path.slice(0, -1);
    }

    if (!path || path === '/') {
      return normalizedUrl.hostname;
    }

    return `${normalizedUrl.hostname}:${path}`;
  }

  public async refresh(): Promise<void> {
    this._clearCache();
    await this._doInitialization();
  }

  private _clearCache(): void {
    try {
      localStorage.removeItem(SiteContextService.CACHE_KEY);
    } catch {
    }
  }
}

export default SiteContextService;
