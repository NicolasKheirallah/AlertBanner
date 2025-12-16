import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { TargetLanguage, ContentType } from '../Alerts/IAlerts';
import { IAlertItem } from "../Alerts/IAlerts";
import { SUPPORTED_LANGUAGES } from '../Utils/AppConstants';
import { logger } from './LoggerService';

export interface ISupportedLanguage {
  code: TargetLanguage;
  name: string;
  nativeName: string;
  flag: string;
  isSupported: boolean;
  columnExists: boolean;
}

export interface ILanguageContent {
  language: TargetLanguage;
  title: string;
  description: string;
  linkDescription?: string;
  availableForAll?: boolean; // If true, this version can be shown to users of other languages
}

export interface IMultiLanguageAlert {
  baseAlert: Omit<IAlertItem, 'title' | 'description' | 'linkDescription'>;
  content: ILanguageContent[];
  languageGroup: string;
}

/**
 * Service for managing language-aware alert content and audience targeting
 */
export class LanguageAwarenessService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private cachedPreferredLanguage: TargetLanguage | null = null;
  private preferredLanguagePromise: Promise<TargetLanguage> | null = null;

  constructor(graphClient: MSGraphClientV3, context: ApplicationCustomizerContext) {
    this.graphClient = graphClient;
    this.context = context;
  }

  /**
   * Get tenant's default language from SharePoint context
   */
  public getTenantDefaultLanguage(): TargetLanguage {
    try {
      const spLanguage = (window as any)._spPageContextInfo?.currentCultureName;
      if (spLanguage) {
        return this.mapLanguageCode(spLanguage.toLowerCase());
      }
      
      const webLanguage = this.context.pageContext.web.language;
      if (webLanguage) {
        return this.mapSharePointLCID(webLanguage);
      }
    } catch (error) {
      logger.warn('LanguageAwarenessService', 'Could not detect tenant language', error);
    }
    
    return TargetLanguage.EnglishUS; // Default fallback
  }

  /**
   * Get all supported languages for the tenant
   */
  public static getSupportedLanguages(): ISupportedLanguage[] {
    return SUPPORTED_LANGUAGES.map(lang => ({
      code: lang.code as TargetLanguage,
      name: lang.name,
      nativeName: lang.nativeName,
      flag: lang.flag,
      isSupported: lang.code === 'en-us', // Only English supported by default
      columnExists: false
    }));
  }

  /**
   * Detect user's preferred language from browser, Azure AD, or SharePoint profile
   */
  public async getUserPreferredLanguage(): Promise<TargetLanguage> {
    if (this.cachedPreferredLanguage) {
      return this.cachedPreferredLanguage;
    }

    if (this.preferredLanguagePromise) {
      return this.preferredLanguagePromise;
    }

    this.preferredLanguagePromise = this.resolveUserPreferredLanguage()
      .then(language => {
        this.cachedPreferredLanguage = language;
        return language;
      })
      .finally(() => {
        this.preferredLanguagePromise = null;
      });

    return this.preferredLanguagePromise;
  }

  private async resolveUserPreferredLanguage(): Promise<TargetLanguage> {
    try {
      const browserLanguage = this.getBrowserLanguage();
      if (browserLanguage) {
        return browserLanguage;
      }

      const userLanguage = await this.getGraphUserLanguage();
      if (userLanguage) {
        return userLanguage;
      }

      const sharePointLanguage = this.getSharePointLanguage();
      if (sharePointLanguage) {
        return sharePointLanguage;
      }

      return this.getTenantDefaultLanguage();
    } catch (error) {
      logger.error('LanguageAwarenessService', 'Error detecting user preferred language', error);
      return TargetLanguage.EnglishUS;
    }
  }

  private getBrowserLanguage(): TargetLanguage | null {
    const browserLanguage = navigator.language?.toLowerCase();
    if (!browserLanguage) {
      return null;
    }

    const mappedLanguage = this.mapLanguageCode(browserLanguage);
    return mappedLanguage || null;
  }

  private async getGraphUserLanguage(): Promise<TargetLanguage | null> {
    try {
      const userProfile = await this.graphClient
        .api('/me')
        .select('preferredLanguage,mailboxSettings')
        .get();

      if (userProfile.preferredLanguage) {
        return this.mapLanguageCode(userProfile.preferredLanguage);
      }
    } catch (error) {
      logger.warn('LanguageAwarenessService', 'Could not retrieve user language from Graph', error);
    }

    return null;
  }

  private getSharePointLanguage(): TargetLanguage | null {
    const spLanguage = (window as any).SPClientContext?.web?.language;
    if (spLanguage) {
      return this.mapSharePointLCID(spLanguage);
    }

    return null;
  }

  /**
   * Map various language codes to our TargetLanguage enum
   */
  private mapLanguageCode(languageCode: string): TargetLanguage {
    const code = languageCode.toLowerCase();
    
    const languageMap: { [key: string]: TargetLanguage } = {
      'en': TargetLanguage.EnglishUS,
      'en-us': TargetLanguage.EnglishUS,
      'en-gb': TargetLanguage.EnglishUS, // Map UK English to US English for now
      'fr': TargetLanguage.FrenchFR,
      'fr-fr': TargetLanguage.FrenchFR,
      'fr-ca': TargetLanguage.FrenchFR, // Map Canadian French to France French
      'de': TargetLanguage.GermanDE,
      'de-de': TargetLanguage.GermanDE,
      'es': TargetLanguage.SpanishES,
      'es-es': TargetLanguage.SpanishES,
      'sv': TargetLanguage.SwedishSE,
      'sv-se': TargetLanguage.SwedishSE,
      'fi': TargetLanguage.FinnishFI,
      'fi-fi': TargetLanguage.FinnishFI,
      'da': TargetLanguage.DanishDK,
      'da-dk': TargetLanguage.DanishDK,
      'nb': TargetLanguage.NorwegianNO,
      'nb-no': TargetLanguage.NorwegianNO,
      'no': TargetLanguage.NorwegianNO
    };

    return languageMap[code] || TargetLanguage.EnglishUS;
  }

  /**
   * Map SharePoint LCID to our TargetLanguage enum
   */
  private mapSharePointLCID(lcid: number): TargetLanguage {
    const lcidMap: { [key: number]: TargetLanguage } = {
      1033: TargetLanguage.EnglishUS,
      1036: TargetLanguage.FrenchFR,
      1031: TargetLanguage.GermanDE,
      1034: TargetLanguage.SpanishES,
      1053: TargetLanguage.SwedishSE,
      1035: TargetLanguage.FinnishFI,
      1030: TargetLanguage.DanishDK,
      1044: TargetLanguage.NorwegianNO
    };

    return lcidMap[lcid] || TargetLanguage.EnglishUS;
  }

  /**
   * Filter and prioritize alerts based on user's preferred language with fallback logic
   */
  public filterAlertsForUser(alerts: IAlertItem[], userLanguage: TargetLanguage): IAlertItem[] {
    const tenantDefault = this.getTenantDefaultLanguage();
    
    // Group alerts by language group
    const alertGroups = new Map<string, IAlertItem[]>();
    const standaloneAlerts: IAlertItem[] = [];
    
    alerts.forEach(alert => {
      if (alert.languageGroup) {
        if (!alertGroups.has(alert.languageGroup)) {
          alertGroups.set(alert.languageGroup, []);
        }
        alertGroups.get(alert.languageGroup)!.push(alert);
      } else {
        // Handle standalone alerts (non-multi-language)
        // Show if: targetLanguage is "all" (case-insensitive), matches user's language, or matches tenant default
        const alertLang = (alert.targetLanguage || TargetLanguage.All)?.toLowerCase();
        const userLang = userLanguage?.toLowerCase();
        const tenantLang = tenantDefault?.toLowerCase();
        
        if (
          alertLang === 'all' || 
          alertLang === userLang ||
          alertLang === tenantLang
        ) {
          standaloneAlerts.push(alert);
        }
      }
    });
    
    // Process language groups with fallback logic
    const selectedAlerts: IAlertItem[] = [];
    
    alertGroups.forEach(groupAlerts => {
      // Try to find alert in user's preferred language
      let selectedAlert = groupAlerts.find(alert => alert.targetLanguage === userLanguage);
      
      // If not found, try to find alert marked as "available for all"
      if (!selectedAlert) {
        const fallbackContent = this.getLanguageContent(groupAlerts, groupAlerts[0].languageGroup!);
        const availableForAllContent = fallbackContent.find(content => content.availableForAll);
        
        if (availableForAllContent) {
          selectedAlert = groupAlerts.find(alert => alert.targetLanguage === availableForAllContent.language);
        }
      }
      
      // If still not found, fall back to tenant default language
      if (!selectedAlert) {
        selectedAlert = groupAlerts.find(alert => alert.targetLanguage === tenantDefault);
      }
      
      // Last resort: pick the first available alert in the group
      if (!selectedAlert) {
        selectedAlert = groupAlerts[0];
      }
      
      if (selectedAlert) {
        selectedAlerts.push(selectedAlert);
      }
    });
    
    return [...selectedAlerts, ...standaloneAlerts];
  }

  /**
   * Create a multi-language alert with content for each language
   */
  public createMultiLanguageAlert(baseAlert: Omit<IAlertItem, 'title' | 'description' | 'linkDescription'>, content: ILanguageContent[]): IMultiLanguageAlert {
    const languageGroup = `lang-group-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;

    return {
      baseAlert: {
        ...baseAlert,
        languageGroup
      },
      content,
      languageGroup
    };
  }

  /**
   * Generate individual alert items from multi-language alert
   */
  public generateAlertItems(multiLangAlert: IMultiLanguageAlert): IAlertItem[] {
    return multiLangAlert.content.map(content => ({
      ...multiLangAlert.baseAlert,
      title: content.title,
      description: content.description,
      linkUrl: multiLangAlert.baseAlert.linkUrl || '',
      linkDescription: content.linkDescription || '',
      targetLanguage: content.language,
      languageGroup: multiLangAlert.languageGroup,
      id: '0' // Will be set by SharePoint when created
    }));
  }

  /**
   * Get language-specific content for editing multi-language alerts
   * Deduplicates by language to ensure each language appears only once
   */
  public getLanguageContent(alerts: IAlertItem[], languageGroup: string): ILanguageContent[] {
    const groupAlerts = alerts.filter(alert => alert.languageGroup === languageGroup);
    const seenLanguages = new Set<string>();
    const uniqueAlerts = groupAlerts.filter(alert => {
      if (seenLanguages.has(alert.targetLanguage)) {
        return false;
      }
      seenLanguages.add(alert.targetLanguage);
      return true;
    });

    return uniqueAlerts.map(alert => ({
      language: alert.targetLanguage,
      title: alert.title,
      description: alert.description,
      linkDescription: alert.linkDescription,
      availableForAll: alert.availableForAll
    }));
  }
}
