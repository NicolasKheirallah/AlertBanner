import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { LIST_NAMES } from "../Utils/AppConstants";
import { logger } from "./LoggerService";
import { SiteIdUtils } from "../Utils/SiteIdUtils";

export class SharePointListLocator {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  
  private listIdCache: Map<string, string> = new Map();
  private graphSiteIdentifierCache: Map<string, string> = new Map();
  private readonly MAX_CACHE_SIZE = 50;

  constructor(graphClient: MSGraphClientV3, context: ApplicationCustomizerContext) {
    this.graphClient = graphClient;
    this.context = context;
  }

  private enforceCacheLimit<T>(cache: Map<string, T>): void {
    if (cache.size >= this.MAX_CACHE_SIZE) {
      cache.clear();
    }
  }

  public getGraphSiteIdentifierFromContext(siteId: string): string {
    const currentUrl = new URL(this.context.pageContext.web.absoluteUrl);
    const siteGuid = SiteIdUtils.normalizeGuid(
      siteId || this.context.pageContext.site.id.toString()
    );
    const webGuid = SiteIdUtils.normalizeGuid(
      this.context.pageContext.web.id.toString()
    );
    return `${currentUrl.hostname},${siteGuid},${webGuid}`;
  }

  public isCurrentSite(siteId: string): boolean {
    if (!siteId) {
      return true;
    }

    const normalized = SiteIdUtils.normalizeGuid(siteId);
    const currentSiteId = SiteIdUtils.normalizeGuid(
      this.context.pageContext.site.id.toString()
    );
    return normalized === currentSiteId;
  }

  public async ensureGraphSiteIdentifier(siteId: string): Promise<string> {
    if (!siteId) {
      return this.getGraphSiteIdentifierFromContext(siteId);
    }

    if (siteId.includes(",")) {
      return siteId;
    }

    if (siteId.startsWith("https://") || siteId.startsWith("http://")) {
      try {
        const siteUrl = new URL(siteId);
        const rawPath = siteUrl.pathname || "";
        const normalizedPath = rawPath.replace(/\/+/g, "/").replace(/\/$/, "");
        return normalizedPath
          ? `${siteUrl.hostname}:${normalizedPath || "/"}`
          : siteUrl.hostname;
      } catch (error) {
        logger.warn(
          "SharePointListLocator",
          "Invalid site URL provided, falling back to context identifier",
          { siteId, error }
        );
        return this.getGraphSiteIdentifierFromContext(siteId);
      }
    }

    if (siteId.includes(":")) {
      return siteId;
    }

    const normalized = SiteIdUtils.normalizeGuid(siteId);
    if (this.graphSiteIdentifierCache.has(normalized)) {
      return this.graphSiteIdentifierCache.get(normalized)!;
    }

    this.enforceCacheLimit(this.graphSiteIdentifierCache);

    if (this.isCurrentSite(siteId)) {
      const identifier = this.getGraphSiteIdentifierFromContext(siteId);
      this.graphSiteIdentifierCache.set(normalized, identifier);
      return identifier;
    }

    try {
      const siteResponse = await this.graphClient
        .api(`/sites/${normalized}`)
        .select("id")
        .get();

      if (siteResponse?.id) {
        this.graphSiteIdentifierCache.set(normalized, siteResponse.id);
        return siteResponse.id;
      }
    } catch (error) {
      const statusCode = (error as any)?.statusCode || (error as any)?.code;
      if (statusCode === 403) {
         logger.info("SharePointListLocator", "Access denied when resolving site identifier, falling back to context value", { siteId });
      } else {
         logger.warn(
            "SharePointListLocator",
            "Unable to resolve graph site identifier, falling back to context derived value",
            { siteId, error }
         );
      }
    }

    const fallback = this.getGraphSiteIdentifierFromContext(siteId);
    this.graphSiteIdentifierCache.set(normalized, fallback);
    return fallback;
  }

  private getListCacheKey(siteIdentifier: string, listTitle: string): string {
    return `${siteIdentifier}|${listTitle.toLowerCase()}`;
  }

  public async resolveListId(
    siteId: string,
    listTitle: string
  ): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const cacheKey = this.getListCacheKey(graphSiteIdentifier, listTitle);
    const cachedId = this.listIdCache.get(cacheKey);
    if (cachedId) {
      return cachedId;
    }

    try {
      // Use OData filter to find the list directly instead of enumerating all lists
      const response = await this.graphClient
        .api(`/sites/${graphSiteIdentifier}/lists`)
        .filter(`displayName eq '${listTitle}'`)
        .select("id,displayName,name")
        .top(1)
        .get();

      if (response.value && response.value.length > 0) {
        const list = response.value[0];
        this.enforceCacheLimit(this.listIdCache);
        this.listIdCache.set(cacheKey, list.id);
        return list.id;
      }
    } catch (error) {
      const statusCode = (error as any)?.statusCode || (error as any)?.code;
      if (statusCode === 403 || (error as any)?.message?.indexOf('Access Denied') > -1) {
        logger.info("SharePointListLocator", `Access denied to list ${listTitle} on ${siteId}. This is expected if the user does not have permission.`);
      } else {
        logger.warn(
            "SharePointListLocator",
            `Unable to resolve list ${listTitle}`,
            { siteId, error }
        );
      }
      throw error;
    }

    const notFoundError = new Error(`LIST_NOT_FOUND:${listTitle}`);
    notFoundError.name = "LIST_NOT_FOUND";
    throw notFoundError;
  }

  public async registerListId(
    siteId: string,
    listTitle: string,
    listId?: string
  ): Promise<void> {
    if (!listId) {
      return;
    }

    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    this.enforceCacheLimit(this.listIdCache);
    this.listIdCache.set(
      this.getListCacheKey(graphSiteIdentifier, listTitle),
      listId
    );
  }

  public async getAlertsListApi(siteId: string): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const listId = await this.resolveListId(siteId, LIST_NAMES.ALERTS);
    return `/sites/${graphSiteIdentifier}/lists/${listId}`;
  }

  public async getAlertTypesListApi(siteId: string): Promise<string> {
    const graphSiteIdentifier = await this.ensureGraphSiteIdentifier(siteId);
    const listId = await this.resolveListId(siteId, LIST_NAMES.ALERT_TYPES);
    return `/sites/${graphSiteIdentifier}/lists/${listId}`;
  }

  private listColumnsCache: Map<string, Set<string>> = new Map();

  /**
   * Get available column names for a list (cached)
   */
  public async getAvailableColumns(alertsListApi: string): Promise<Set<string>> {
    if (this.listColumnsCache.has(alertsListApi)) {
      return this.listColumnsCache.get(alertsListApi)!;
    }

    try {
      const columnsResponse = await this.graphClient
        .api(`${alertsListApi}/columns?$select=name`)
        .get();

      const availableColumns = new Set<string>(
        (columnsResponse.value || [])
          .map((c: any) => (c.name || "").trim())
          .filter(Boolean)
      );

      this.enforceCacheLimit(this.listColumnsCache);
      this.listColumnsCache.set(alertsListApi, availableColumns);
      return availableColumns;
    } catch (error) {
      logger.warn(
        "SharePointListLocator",
        "Could not retrieve column metadata, falling back to minimal column set",
        { list: alertsListApi, error }
      );
      const fallback = new Set<string>([
        "Title",
        "Created",
        "Modified",
        "Author",
        "Editor",
        "Attachments",
      ]);
      this.enforceCacheLimit(this.listColumnsCache);
      this.listColumnsCache.set(alertsListApi, fallback);
      return fallback;
    }
  }

  public async getSiteUrlFromIdentifier(siteId: string): Promise<string> {
    if (!siteId) {
      return this.context.pageContext.web.absoluteUrl;
    }

    if (siteId.startsWith("https://")) {
      return siteId;
    }

    if (siteId.includes(":") && !siteId.includes(",")) {
      const [hostname, path = "/"] = siteId.split(":");
      return `https://${hostname}${path === "/" ? "" : path}`;
    }

    if (siteId.includes(",")) {
      try {
        const siteDetails = await this.graphClient
          .api(`/sites/${siteId}`)
          .select("webUrl")
          .get();

        if (siteDetails?.webUrl) {
          return siteDetails.webUrl;
        }
      } catch (error) {
        const statusCode = (error as any)?.statusCode || (error as any)?.code;
        if (statusCode === 403) {
            logger.info("SharePointListLocator", "Access denied when resolving site URL, using fallback", { siteId });
        } else {
            logger.warn(
            "SharePointListLocator",
            "Unable to resolve site URL from identifier",
            { siteId, error }
            );
        }
      }
    }

    return this.context.pageContext.web.absoluteUrl;
  }
}
