import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { logger } from './LoggerService';

export enum GraphPermission {
  SitesReadAll = 'Sites.Read.All',
  SitesReadWriteAll = 'Sites.ReadWrite.All',
  UserRead = 'User.Read',
  MailSend = 'Mail.Send',
  GroupMemberReadAll = 'GroupMember.Read.All',
  DirectoryReadAll = 'Directory.Read.All'
}

export interface IPermissionStatus {
  scope: GraphPermission;
  granted: boolean;
  error?: string;
}

export interface IAuditableOperation {
  operation: string;
  targetSite?: string;
  targetSiteUrl?: string;
  targetSiteName?: string;
  targetList?: string;
  itemCount?: number;
  justification?: string;
  error?: string;
}

export class PermissionService {
  private static instance: PermissionService;
  private context: ApplicationCustomizerContext;
  private graphClient: MSGraphClientV3 | undefined;
  private permissionCache: Map<GraphPermission, IPermissionStatus> = new Map();
  private lastPermissionCheck: number = 0;
  private readonly PERMISSION_CACHE_TTL = 5 * 60 * 1000; // 5 minutes

  private constructor(context: ApplicationCustomizerContext) {
    this.context = context;
  }

  public static getInstance(context: ApplicationCustomizerContext): PermissionService {
    if (!PermissionService.instance) {
      PermissionService.instance = new PermissionService(context);
    }
    return PermissionService.instance;
  }

  public async initialize(): Promise<void> {
    try {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3') as MSGraphClientV3;
      await this.validateAllPermissions();
    } catch (error) {
      logger.error('PermissionService', 'Failed to initialize Graph client', error);
      throw error;
    }
  }

  private async ensureGraphClient(): Promise<MSGraphClientV3 | undefined> {
    if (this.graphClient) {
      return this.graphClient;
    }

    try {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3') as MSGraphClientV3;
      return this.graphClient;
    } catch (error) {
      logger.warn('PermissionService', 'Unable to acquire Graph client', error);
      return undefined;
    }
  }

  public async validateAllPermissions(): Promise<IPermissionStatus[]> {
    const now = Date.now();
    if (now - this.lastPermissionCheck < this.PERMISSION_CACHE_TTL && this.permissionCache.size > 0) {
      return Array.from(this.permissionCache.values());
    }

    const permissions = [
      GraphPermission.SitesReadWriteAll,
      GraphPermission.UserRead,
      GraphPermission.MailSend,
      GraphPermission.GroupMemberReadAll,
      GraphPermission.DirectoryReadAll
    ];

    const results = await Promise.all(
      permissions.map(scope => this.checkPermission(scope))
    );

    this.lastPermissionCheck = now;
    return results;
  }

  private async checkPermission(scope: GraphPermission): Promise<IPermissionStatus> {
    const graphClient = await this.ensureGraphClient();
    if (!graphClient) {
      return { scope, granted: false, error: 'Graph client not initialized' };
    }

    try {
      switch (scope) {
        case GraphPermission.SitesReadWriteAll:
          await graphClient.api('/sites/root').select('id').get();
          break;
        case GraphPermission.MailSend:
          // Mail.Send can only be tested by attempting to send
          break;
        case GraphPermission.UserRead:
          await graphClient.api('/me').select('id').get();
          break;
        case GraphPermission.GroupMemberReadAll:
          await graphClient.api('/me/memberOf').top(1).get();
          break;
        case GraphPermission.DirectoryReadAll:
          await graphClient.api('/organization').top(1).get();
          break;
      }

      const status: IPermissionStatus = { scope, granted: true };
      this.permissionCache.set(scope, status);
      return status;
    } catch (error: any) {
      const is403 = error.statusCode === 403 || error.code === 'Forbidden';
      const is401 = error.statusCode === 401 || error.code === 'Unauthorized';
      
      const status: IPermissionStatus = {
        scope,
        granted: false,
        error: is403 ? 'Permission not granted' : is401 ? 'Admin consent required' : error.message
      };
      
      this.permissionCache.set(scope, status);
      
      if (is401 || is403) {
        logger.warn('PermissionService', `Permission ${scope} not available`, {
          error: status.error,
          requiresAdminConsent: is401
        });
      }
      
      return status;
    }
  }

  public async canWriteSites(): Promise<boolean> {
    const status = await this.checkPermission(GraphPermission.SitesReadWriteAll);
    return status.granted;
  }

  public async canSendMail(): Promise<boolean> {
    const status = await this.checkPermission(GraphPermission.MailSend);
    return status.granted;
  }

  // Log auditable operation for security compliance - all write operations must be logged
  public logAuditableOperation(operation: { operation: string; targetSite?: string; targetSiteUrl?: string; targetSiteName?: string; targetList?: string; itemCount?: number; justification?: string; error?: string; }): void {
    const auditEntry = {
      timestamp: new Date().toISOString(),
      userId: this.context.pageContext.user.loginName,
      userDisplayName: this.context.pageContext.user.displayName,
      userEmail: this.context.pageContext.user.email,
      correlationId: this.getCorrelationId(),
      currentSite: this.context.pageContext.web.absoluteUrl,
      ...operation
    };

    logger.info('AUDIT', `Auditable operation: ${operation.operation}`, auditEntry);
    this.sendToAuditLog(auditEntry);
  }

  public async executeWriteOperation<T>(
    operation: () => Promise<T>,
    audit: IAuditableOperation
  ): Promise<T> {
    const enrichedAudit = {
      ...audit,
      targetSiteUrl: audit.targetSiteUrl || this.context.pageContext.web.absoluteUrl,
      targetSiteName: audit.targetSiteName || this.context.pageContext.web.title
    };

    const canWrite = await this.canWriteSites();
    if (!canWrite) {
      const siteContext = enrichedAudit.targetSiteName || enrichedAudit.targetSiteUrl || 'target site';
      const error = new Error(
        `Sites.ReadWrite.All permission not granted for ${siteContext}. ` +
        'Please request admin consent for this permission in the SharePoint Admin Center.'
      );
      logger.error('PermissionService', 'Write operation blocked - permission not granted', {
        operation: audit.operation,
        targetSite: audit.targetSite,
        targetSiteUrl: enrichedAudit.targetSiteUrl,
        targetSiteName: enrichedAudit.targetSiteName,
        user: this.context.pageContext.user.loginName
      });
      throw error;
    }

    this.logAuditableOperation({
      ...enrichedAudit,
      operation: `${audit.operation}_START`
    });

    try {
      const result = await operation();
      
      this.logAuditableOperation({
        ...enrichedAudit,
        operation: `${audit.operation}_SUCCESS`
      });
      
      return result;
    } catch (error) {
      const errorMessage = (error as Error).message;
      const siteContext = enrichedAudit.targetSiteName || enrichedAudit.targetSiteUrl || 'target site';
      const enrichedError = new Error(
        `Operation ${audit.operation} failed on site "${siteContext}": ${errorMessage}`
      );
      
      this.logAuditableOperation({
        ...enrichedAudit,
        operation: `${audit.operation}_FAILED`,
        error: enrichedError.message
      });
      
      throw enrichedError;
    }
  }

  public getAdminConsentUrl(): string {
    const tenantId = this.context.pageContext.aadInfo?.tenantId;
    const clientId = this.context.pageContext.aadInfo?.instanceId || 
                     '00000003-0000-0ff1-ce00-000000000000';
    
    const scopes = [
      GraphPermission.SitesReadWriteAll,
      GraphPermission.MailSend
    ].join(' ');

    return `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${clientId}&scope=${encodeURIComponent(scopes)}`;
  }

  private getCorrelationId(): string {
    return (this.context as any).correlationId || 
           `${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
  }

  private sendToAuditLog(entry: any): void {
    if ((window as any).appInsights) {
      (window as any).appInsights.trackEvent({
        name: 'AlertBannerAudit',
        properties: entry
      });
    }
  }

  public clearCache(): void {
    this.permissionCache.clear();
    this.lastPermissionCheck = 0;
    logger.info('PermissionService', 'Permission cache cleared');
  }
}

export default PermissionService;
