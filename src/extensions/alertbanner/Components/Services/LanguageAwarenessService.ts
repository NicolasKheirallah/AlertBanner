import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { Guid } from "@microsoft/sp-core-library";
import {
  TargetLanguage,
  TranslationStatus,
  IAlertItem,
  ILanguageContent,
  IMultiLanguageAlert,
} from "../Alerts/IAlerts";
import { SUPPORTED_LANGUAGES } from "../Utils/AppConstants";
import { logger } from "./LoggerService";
import {
  ILanguagePolicy,
  normalizeLanguagePolicy,
} from "./LanguagePolicyService";

// Re-export ILanguageContent for convenience
export type { ILanguageContent } from "../Alerts/IAlerts";

export interface ISupportedLanguage {
  code: TargetLanguage;
  name: string;
  nativeName: string;
  flag: string;
  isSupported: boolean;
  columnExists: boolean;
}

export interface ILanguageDetectionResult {
  language: TargetLanguage;
  source: 'user-override' | 'browser' | 'graph' | 'sharepoint' | 'tenant-default' | 'fallback';
}

export class LanguageAwarenessService {
  private graphClient: MSGraphClientV3;
  private context: ApplicationCustomizerContext;
  private cachedPreferredLanguage: TargetLanguage | null = null;
  private preferredLanguagePromise: Promise<TargetLanguage> | null = null;
  private static readonly STORAGE_KEY = 'AlertBanner_UserPreferredLanguage';

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
  ) {
    this.graphClient = graphClient;
    this.context = context;
  }

  public getTenantDefaultLanguage(): TargetLanguage {
    try {
      const spLanguage = (window as any)._spPageContextInfo?.currentCultureName;
      if (spLanguage) {
        const mapped = this.mapLanguageCode(spLanguage.toLowerCase());
        if (mapped) return mapped;
      }

      const webLanguage = this.context.pageContext.web.language;
      if (webLanguage) {
        const mapped = this.mapSharePointLCID(webLanguage);
        if (mapped) return mapped;
      }
    } catch (error) {
      logger.warn(
        "LanguageAwarenessService",
        "Could not detect tenant language",
        error,
      );
    }

    return TargetLanguage.EnglishUS;
  }

  public static getSupportedLanguages(): ISupportedLanguage[] {
    return SUPPORTED_LANGUAGES.map((lang) => ({
      code: lang.code as TargetLanguage,
      name: lang.name,
      nativeName: lang.nativeName,
      flag: lang.flag,
      isSupported: lang.code === "en-us", // Only English supported by default
      columnExists: false,
    }));
  }

  public async getUserPreferredLanguage(): Promise<TargetLanguage> {
    if (this.cachedPreferredLanguage) {
      return this.cachedPreferredLanguage;
    }

    if (this.preferredLanguagePromise) {
      return this.preferredLanguagePromise;
    }

    this.preferredLanguagePromise = this.resolveUserPreferredLanguage()
      .then((result) => {
        this.cachedPreferredLanguage = result.language;
        return result.language;
      })
      .finally(() => {
        this.preferredLanguagePromise = null;
      });

    return this.preferredLanguagePromise;
  }

  /**
   * Get detailed language detection info including source
   */
  public async getUserPreferredLanguageWithSource(): Promise<ILanguageDetectionResult> {
    const language = await this.getUserPreferredLanguage();
    
    // Check if it was user override
    const stored = this.getStoredLanguagePreference();
    if (stored) {
      return { language, source: 'user-override' };
    }
    
    // Re-detect to find source
    const result = await this.resolveUserPreferredLanguage();
    return result;
  }

  /**
   * Allow user to manually set their preferred language
   */
  public setUserPreferredLanguage(language: TargetLanguage): void {
    this.cachedPreferredLanguage = language;
    this.storeLanguagePreference(language);
    logger.info("LanguageAwarenessService", `User manually set language to ${language}`);
  }

  /**
   * Clear stored language preference and re-detect
   */
  public clearUserLanguagePreference(): void {
    this.cachedPreferredLanguage = null;
    this.clearStoredLanguagePreference();
    logger.info("LanguageAwarenessService", "Cleared user language preference, will re-detect");
  }

  private getStoredLanguagePreference(): TargetLanguage | null {
    try {
      const stored = localStorage.getItem(LanguageAwarenessService.STORAGE_KEY);
      if (stored && Object.values(TargetLanguage).includes(stored as TargetLanguage)) {
        return stored as TargetLanguage;
      }
    } catch {
      // localStorage not available
    }
    return null;
  }

  private storeLanguagePreference(language: TargetLanguage): void {
    try {
      localStorage.setItem(LanguageAwarenessService.STORAGE_KEY, language);
    } catch {
      // localStorage not available
    }
  }

  private clearStoredLanguagePreference(): void {
    try {
      localStorage.removeItem(LanguageAwarenessService.STORAGE_KEY);
    } catch {
      // localStorage not available
    }
  }

  private async resolveUserPreferredLanguage(): Promise<ILanguageDetectionResult> {
    try {
      // 1. Check user override first (stored preference)
      const storedPreference = this.getStoredLanguagePreference();
      if (storedPreference) {
        logger.debug("LanguageAwarenessService", "Using stored user preference", storedPreference);
        return { language: storedPreference, source: 'user-override' };
      }

      // 2. Check ALL browser languages (not just first)
      const browserLanguages = this.getAllBrowserLanguages();
      if (browserLanguages.length > 0) {
        logger.debug("LanguageAwarenessService", "Using browser language", browserLanguages[0]);
        return { language: browserLanguages[0], source: 'browser' };
      }

      // 3. Fall back to Graph API user profile
      const userLanguage = await this.getGraphUserLanguage();
      if (userLanguage) {
        logger.debug("LanguageAwarenessService", "Using Graph API language", userLanguage);
        return { language: userLanguage, source: 'graph' };
      }

      // 4. Fall back to SharePoint context
      const sharePointLanguage = this.getSharePointLanguage();
      if (sharePointLanguage) {
        logger.debug("LanguageAwarenessService", "Using SharePoint language", sharePointLanguage);
        return { language: sharePointLanguage, source: 'sharepoint' };
      }

      // 5. Final fallback to tenant default
      const tenantDefault = this.getTenantDefaultLanguage();
      logger.debug("LanguageAwarenessService", "Using tenant default language", tenantDefault);
      return { language: tenantDefault, source: 'tenant-default' };
    } catch (error) {
      logger.error(
        "LanguageAwarenessService",
        "Error detecting user preferred language",
        error,
      );
      return { language: TargetLanguage.EnglishUS, source: 'fallback' };
    }
  }

  /**
   * Get the first matching browser language
   */
  private getBrowserLanguage(): TargetLanguage | null {
    const languages = this.getAllBrowserLanguages();
    return languages.length > 0 ? languages[0] : null;
  }

  /**
   * Get ALL supported browser languages in order of preference
   */
  private getAllBrowserLanguages(): TargetLanguage[] {
    const languages = navigator.languages || [navigator.language];
    const supportedLanguages: TargetLanguage[] = [];

    for (const lang of languages) {
      if (lang) {
        const mappedLanguage = this.mapLanguageCode(lang.toLowerCase());
        if (mappedLanguage && !supportedLanguages.includes(mappedLanguage)) {
          supportedLanguages.push(mappedLanguage);
        }
      }
    }

    return supportedLanguages;
  }

  private async getGraphUserLanguage(): Promise<TargetLanguage | null> {
    try {
      const userProfile = await this.graphClient
        .api("/me")
        .select("preferredLanguage")
        .get();

      if (userProfile.preferredLanguage) {
        return this.mapLanguageCode(userProfile.preferredLanguage);
      }
    } catch (error) {
      logger.warn(
        "LanguageAwarenessService",
        "Could not retrieve user language from Graph",
        error,
      );
    }

    return null;
  }

  private getSharePointLanguage(): TargetLanguage | null {
    const contextLanguage =
      this.context?.pageContext?.cultureInfo?.currentUICultureName;
    if (contextLanguage) {
      return this.mapLanguageCode(contextLanguage.toLowerCase());
    }

    const spLanguage = (window as any).SPClientContext?.web?.language;
    if (spLanguage) {
      return this.mapSharePointLCID(spLanguage);
    }

    return null;
  }

  private mapLanguageCode(languageCode: string): TargetLanguage | null {
    const code = languageCode.toLowerCase();

    const languageMap: { [key: string]: TargetLanguage } = {
      en: TargetLanguage.EnglishUS,
      "en-us": TargetLanguage.EnglishUS,
      "en-gb": TargetLanguage.EnglishUS, // Map UK English to US English for now
      fr: TargetLanguage.FrenchFR,
      "fr-fr": TargetLanguage.FrenchFR,
      "fr-ca": TargetLanguage.FrenchFR, // Map Canadian French to France French
      de: TargetLanguage.GermanDE,
      "de-de": TargetLanguage.GermanDE,
      es: TargetLanguage.SpanishES,
      "es-es": TargetLanguage.SpanishES,
      sv: TargetLanguage.SwedishSE,
      "sv-se": TargetLanguage.SwedishSE,
      fi: TargetLanguage.FinnishFI,
      "fi-fi": TargetLanguage.FinnishFI,
      da: TargetLanguage.DanishDK,
      "da-dk": TargetLanguage.DanishDK,
      nb: TargetLanguage.NorwegianNO,
      "nb-no": TargetLanguage.NorwegianNO,
      no: TargetLanguage.NorwegianNO,
    };

    return languageMap[code] || null;
  }

  private mapSharePointLCID(lcid: number): TargetLanguage | null {
    const lcidMap: { [key: number]: TargetLanguage } = {
      1033: TargetLanguage.EnglishUS,
      1036: TargetLanguage.FrenchFR,
      1031: TargetLanguage.GermanDE,
      1034: TargetLanguage.SpanishES,
      1053: TargetLanguage.SwedishSE,
      1035: TargetLanguage.FinnishFI,
      1030: TargetLanguage.DanishDK,
      1044: TargetLanguage.NorwegianNO,
    };

    return lcidMap[lcid] || null;
  }

  /**
   * Get user's full language preference list in priority order
   */
  public getUserLanguagePreferences(): TargetLanguage[] {
    const preferences: TargetLanguage[] = [];
    
    // 1. Stored user override
    const stored = this.getStoredLanguagePreference();
    if (stored) {
      preferences.push(stored);
    }
    
    // 2. All supported browser languages
    const browserLanguages = this.getAllBrowserLanguages();
    for (const lang of browserLanguages) {
      if (!preferences.includes(lang)) {
        preferences.push(lang);
      }
    }
    
    // 3. Tenant default (always add as final fallback)
    const tenantDefault = this.getTenantDefaultLanguage();
    if (!preferences.includes(tenantDefault)) {
      preferences.push(tenantDefault);
    }
    
    return preferences;
  }

  public filterAlertsForUser(
    alerts: IAlertItem[],
    userLanguage: TargetLanguage,
    policy?: ILanguagePolicy,
  ): IAlertItem[] {
    const tenantDefault = this.getTenantDefaultLanguage();
    const effectivePolicy = normalizeLanguagePolicy(policy);
    
    // Get full preference list for better matching
    const userPreferences = this.getUserLanguagePreferences();

    // Group alerts by language group
    const alertGroups = new Map<string, IAlertItem[]>();
    const standaloneAlerts: IAlertItem[] = [];

    alerts.forEach((alert) => {
      if (alert.languageGroup) {
        if (!alertGroups.has(alert.languageGroup)) {
          alertGroups.set(alert.languageGroup, []);
        }
        alertGroups.get(alert.languageGroup)!.push(alert);
      } else {
        const alertLang = (
          alert.targetLanguage || TargetLanguage.All
        )?.toLowerCase();
        const userLang = userLanguage?.toLowerCase();
        const tenantLang = tenantDefault?.toLowerCase();

        if (
          alertLang === "all" ||
          alertLang === userLang ||
          alertLang === tenantLang
        ) {
          standaloneAlerts.push(alert);
        }
      }
    });

    // Process language groups with fallback logic
    const selectedAlerts: IAlertItem[] = [];

    alertGroups.forEach((groupAlerts) => {
      let candidateAlerts = groupAlerts;
      if (
        effectivePolicy.workflow.enabled &&
        effectivePolicy.workflow.requireApprovedForDisplay
      ) {
        candidateAlerts = groupAlerts.filter(
          (alert) =>
            (alert.translationStatus || TranslationStatus.Approved) ===
            TranslationStatus.Approved,
        );
        if (candidateAlerts.length === 0) {
          return;
        }
      }

      // Try user's full preference list in order
      let selectedAlert: IAlertItem | undefined;
      
      for (const preferredLang of userPreferences) {
        selectedAlert = candidateAlerts.find(
          (alert) => alert.targetLanguage === preferredLang,
        );
        if (selectedAlert) {
          logger.debug(
            "LanguageAwarenessService",
            `Matched alert to user's ${preferredLang} preference`,
            { languageGroup: groupAlerts[0]?.languageGroup }
          );
          break;
        }
      }

      // Fall back to availableForAll
      if (!selectedAlert) {
        selectedAlert = candidateAlerts.find(
          (alert) => alert.availableForAll,
        );
      }

      // Final fallback: first available
      if (!selectedAlert) {
        selectedAlert = candidateAlerts[0];
        logger.debug(
          "LanguageAwarenessService",
          `No preferred language match, using first available`,
          { 
            languageGroup: groupAlerts[0]?.languageGroup,
            selectedLanguage: selectedAlert?.targetLanguage 
          }
        );
      }

      if (selectedAlert) {
        selectedAlerts.push(
          this.applyFieldInheritance(
            selectedAlert,
            candidateAlerts,
            effectivePolicy,
            tenantDefault,
          ),
        );
      }
    });

    return [...selectedAlerts, ...standaloneAlerts];
  }

  public createMultiLanguageAlert(
    baseAlert: Omit<IAlertItem, "title" | "description" | "linkDescription">,
    content: ILanguageContent[],
  ): IMultiLanguageAlert {
    const languageGroup =
      baseAlert.languageGroup || `lang-group-${Guid.newGuid().toString()}`;

    return {
      baseAlert: {
        ...baseAlert,
        languageGroup,
      },
      content,
      languageGroup,
    };
  }

  public validateMultiLanguageContent(content: ILanguageContent[]): {
    isValid: boolean;
    errors: string[];
  } {
    const errors: string[] = [];

    if (content.length === 0) {
      errors.push("At least one language must be added");
      return { isValid: false, errors };
    }

    let hasCompleteLanguage = false;

    content.forEach((langContent) => {
      const langErrors: string[] = [];

      if (!langContent.title || langContent.title.trim().length < 3) {
        langErrors.push(`Title is required (min 3 characters)`);
      }

      if (
        !langContent.description ||
        langContent.description.trim().length < 10
      ) {
        langErrors.push(`Description is required (min 10 characters)`);
      }

      if (langErrors.length === 0) {
        hasCompleteLanguage = true;
      }
    });

    if (!hasCompleteLanguage) {
      errors.push(
        "At least one language must have complete content (title and description)",
      );
    }

    return { isValid: errors.length === 0, errors };
  }

  public generateAlertItems(multiLangAlert: IMultiLanguageAlert): IAlertItem[] {
    return multiLangAlert.content.map((content) => ({
      ...multiLangAlert.baseAlert,
      title: content.title,
      description: content.description,
      linkUrl: multiLangAlert.baseAlert.linkUrl || "",
      linkDescription: content.linkDescription || "",
      targetLanguage: content.language,
      languageGroup: multiLangAlert.languageGroup,
      translationStatus:
        content.translationStatus || TranslationStatus.Approved,
      id: "0", // Will be set by SharePoint when created
    }));
  }

  public getLanguageContent(
    alerts: IAlertItem[],
    languageGroup: string,
  ): ILanguageContent[] {
    const groupAlerts = alerts.filter(
      (alert) => alert.languageGroup === languageGroup,
    );
    const seenLanguages = new Set<string>();
    const uniqueAlerts = groupAlerts.filter((alert) => {
      if (seenLanguages.has(alert.targetLanguage)) {
        return false;
      }
      seenLanguages.add(alert.targetLanguage);
      return true;
    });

    return uniqueAlerts.map((alert) => ({
      language: alert.targetLanguage,
      title: alert.title,
      description: alert.description,
      linkDescription: alert.linkDescription,
      availableForAll: alert.availableForAll,
      translationStatus: alert.translationStatus || TranslationStatus.Approved,
    }));
  }

  private applyFieldInheritance(
    selectedAlert: IAlertItem,
    groupAlerts: IAlertItem[],
    policy: ILanguagePolicy,
    tenantDefault: TargetLanguage,
  ): IAlertItem {
    if (!policy.inheritance.enabled) {
      return selectedAlert;
    }

    const fallbackLanguage =
      policy.fallbackLanguage === "tenant-default"
        ? tenantDefault
        : policy.fallbackLanguage;
    const fallbackAlert =
      groupAlerts.find((alert) => alert.targetLanguage === fallbackLanguage) ||
      groupAlerts.find((alert) => alert.availableForAll) ||
      groupAlerts[0];

    if (!fallbackAlert) {
      return selectedAlert;
    }

    const merged: IAlertItem = { ...selectedAlert };

    if (policy.inheritance.fields.title && !merged.title?.trim()) {
      merged.title = fallbackAlert.title;
    }
    if (policy.inheritance.fields.description && !merged.description?.trim()) {
      merged.description = fallbackAlert.description;
    }
    if (
      policy.inheritance.fields.linkDescription &&
      !merged.linkDescription?.trim()
    ) {
      merged.linkDescription = fallbackAlert.linkDescription;
    }

    return merged;
  }
}
