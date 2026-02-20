import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SharePointListLocator } from "./SharePointListLocator";
import {
  ListProvisioningService,
  IRepairResult,
} from "./ListProvisioningService";
import { AlertOperationsService } from "./AlertOperationsService";
import { logger } from "./LoggerService";
import { AlertFilters } from "../Utils/AlertFilters";
import {
  IAlertItem,
  IAlertType,
  ContentType,
  AlertPriority,
} from "../Alerts/IAlerts";
import { ILanguagePolicy } from "./LanguagePolicyService";
import { LIST_NAMES } from "../Utils/AppConstants";

export type { IRepairResult }; // Re-export for consumers

export class SharePointAlertService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private locator: SharePointListLocator;
  private provisioning: ListProvisioningService;
  private operations: AlertOperationsService;

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
  ) {
    this.graphClient = graphClient;
    this.context = context;
    this.locator = new SharePointListLocator(graphClient, context);
    this.provisioning = new ListProvisioningService(
      graphClient,
      context,
      this.locator,
    );
    this.operations = new AlertOperationsService(
      graphClient,
      context,
      this.locator,
    );
  }

  public async checkListsNeeded() {
    return this.provisioning.checkListsNeeded();
  }
  public async initializeLists(siteId?: string) {
    return this.provisioning.initializeLists(siteId);
  }
  public async repairAlertsList(
    siteId: string,
    cb?: (msg: string, percent: number) => void,
  ) {
    return this.provisioning.repairAlertsList(siteId, cb);
  }
  public async getSupportedLanguages(siteId?: string) {
    return this.provisioning.getSupportedLanguages(siteId);
  }
  public async addLanguageSupport(lang: string, siteId?: string) {
    return this.provisioning.addLanguageSupport(lang, siteId);
  }
  public async removeLanguageSupport(lang: string) {
    return this.provisioning.removeLanguageSupport(lang);
  }
  public async updateSupportedLanguages(siteId: string, langs: string[]) {
    return this.provisioning.updateSupportedLanguages(siteId, langs);
  }

  public async getActiveAlerts(siteId: string) {
    return this.operations.getActiveAlerts(siteId);
  }
  public async createAlert(alert: Record<string, unknown>) {
    return this.operations.createAlert(alert);
  }
  public async updateAlert(id: string, updates: Record<string, unknown>) {
    return this.operations.updateAlert(id, updates);
  }
  public async deleteAlert(id: string) {
    return this.operations.deleteAlert(id);
  }
  public async deleteAlerts(ids: string[]) {
    return this.operations.deleteAlerts(ids);
  }
  public async updateAlertsSortOrder(
    alertOrders: { alertId: string; sortOrder: number }[],
  ) {
    return this.operations.updateAlertsSortOrder(alertOrders);
  }
  public async getAlertTypes(siteId?: string) {
    return this.operations.getAlertTypes(siteId);
  }
  public async saveAlertTypes(alertTypes: IAlertType[]) {
    return this.operations.saveAlertTypes(alertTypes);
  }
  public async getDraftAlerts(siteId: string) {
    return this.operations.getDraftAlerts(siteId);
  }
  public async saveDraft(draft: Record<string, unknown>) {
    return this.operations.saveDraft(draft);
  }
  public async deleteDraft(id: string) {
    return this.operations.deleteDraft(id);
  }
  public async addAttachment(
    listId: string,
    itemId: number,
    name: string,
    content: ArrayBuffer,
    siteId?: string,
  ) {
    return this.operations.addAttachment(listId, itemId, name, content, siteId);
  }
  public async deleteAttachment(
    listId: string,
    itemId: number,
    name: string,
    siteId?: string,
  ) {
    return this.operations.deleteAttachment(listId, itemId, name, siteId);
  }
  public async getTemplateAlerts(siteId: string) {
    return this.operations.getTemplateAlerts(siteId);
  }
  public async getLanguagePolicy(siteId?: string) {
    return this.operations.getLanguagePolicy(siteId);
  }
  public async saveLanguagePolicy(policy: ILanguagePolicy, siteId?: string) {
    return this.operations.saveLanguagePolicy(policy, siteId);
  }

  public async submitAlert(id: string, reviewerId?: string) {
    return this.operations.submitAlert(id, reviewerId);
  }
  public async approveAlert(id: string, comments?: string) {
    return this.operations.approveAlert(id, comments);
  }
  public async rejectAlert(id: string, comments?: string) {
    return this.operations.rejectAlert(id, comments);
  }

  public getCurrentSiteId() {
    return this.context.pageContext.site.id.toString();
  }

  public async getAlertsListId(siteId?: string): Promise<string> {
    const targetSite = siteId || this.getCurrentSiteId();
    return this.locator.resolveListId(targetSite, LIST_NAMES.ALERTS);
  }

  public parseAlertId(alertId: string): { siteId: string; itemId: string } {
    return this.operations.parseAlertId(alertId);
  }

  public getAlertSiteId(alertId: string): string {
    return this.operations.parseAlertId(alertId).siteId;
  }

  public async getAlerts(siteIds?: string[]): Promise<IAlertItem[]> {
    const sites = await this.resolveSiteIds(siteIds);
    const allAlerts: IAlertItem[] = [];
    const batchSize = 2;

    for (let i = 0; i < sites.length; i += batchSize) {
      const batch = sites.slice(i, i + batchSize);
      const results = await Promise.allSettled(
        batch.map((s) => this.operations.getAlertsForSite(s)),
      );
      results.forEach((r) => {
        if (r.status === "fulfilled") allAlerts.push(...r.value);
        else
          logger.warn(
            "SharePointAlertService",
            "Failed to get alerts batch",
            r.reason,
          );
      });
    }

    const unique = AlertFilters.removeDuplicates(allAlerts);
    return unique.sort(
      (a, b) =>
        new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime(),
    );
  }

  public async getAlertsAndTemplates(
    siteIds?: string[],
  ): Promise<IAlertItem[]> {
    const sites = await this.resolveSiteIds(siteIds);
    const allItems: IAlertItem[] = [];
    const batchSize = 2;

    for (let i = 0; i < sites.length; i += batchSize) {
      const batch = sites.slice(i, i + batchSize);
      const results = await Promise.allSettled(
        batch.map((siteId) => this.operations.getManageItemsForSite(siteId)),
      );

      results.forEach((r) => {
        if (r.status === "fulfilled") allItems.push(...r.value);
        else
          logger.warn(
            "SharePointAlertService",
            "Failed to get alerts/templates batch",
            r.reason,
          );
      });
    }

    const unique = AlertFilters.removeDuplicates(allItems);
    return unique.sort(
      (a, b) =>
        new Date(b.createdDate).getTime() - new Date(a.createdDate).getTime(),
    );
  }

  private async resolveSiteIds(siteIds?: string[]): Promise<string[]> {
    if (siteIds) return siteIds;

    try {
      const { SiteContextService } = await import("./SiteContextService");
      const siteContext = SiteContextService.getInstance(
        this.context,
        this.graphClient,
      );
      await siteContext.initialize();
      const hierarchy = siteContext.getAlertSourceSites();

      const unique = new Set<string>();
      hierarchy.forEach((s) =>
        unique.add(
          s.includes(",")
            ? s.split(",")[1]
            : s.replace(/[{}]/g, "").toLowerCase(),
        ),
      ); // Simplified logic
      // I will trust the input for now or use Array.from(unique) but mapping back to original strings?
      return hierarchy;
    } catch (e) {
      return [this.context.pageContext.site.id.toString()];
    }
  }
}
