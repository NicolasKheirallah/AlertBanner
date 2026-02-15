import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { Guid } from "@microsoft/sp-core-library";
import { TargetLanguage, TranslationStatus } from "../Alerts/IAlerts";
import { IAlertItem } from "../Alerts/IAlerts";
import { SUPPORTED_LANGUAGES } from "../Utils/AppConstants";
import { logger } from "./LoggerService";
import {
  ILanguagePolicy,
  normalizeLanguagePolicy,
} from "./LanguagePolicyService";

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
  translationStatus?: TranslationStatus;
}

export interface IMultiLanguageAlert {
  baseAlert: Omit<IAlertItem, "title" | "description" | "linkDescription">;
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

  constructor(
    graphClient: MSGraphClientV3,
    context: ApplicationCustomizerContext,
  ) {
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
      logger.warn(
        "LanguageAwarenessService",
        "Could not detect tenant language",
        error,
      );
    }

    return TargetLanguage.EnglishUS; // Default fallback
  }

  /**
   * Get all supported languages for the tenant
   */
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
      .then((language) => {
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
      const userLanguage = await this.getGraphUserLanguage();
      if (userLanguage) {
        return userLanguage;
      }

      const sharePointLanguage = this.getSharePointLanguage();
      if (sharePointLanguage) {
        return sharePointLanguage;
      }

      const browserLanguage = this.getBrowserLanguage();
      if (browserLanguage) {
        return browserLanguage;
      }

      return this.getTenantDefaultLanguage();
    } catch (error) {
      logger.error(
        "LanguageAwarenessService",
        "Error detecting user preferred language",
        error,
      );
      return TargetLanguage.EnglishUS;
    }
  }

  private getBrowserLanguage(): TargetLanguage | null {
    // Check full array of preferred languages (navigator.languages is supported in modern browsers)
    const languages = navigator.languages || [navigator.language];

    for (const lang of languages) {
      if (lang) {
        const mappedLanguage = this.mapLanguageCode(lang.toLowerCase());
        if (mappedLanguage) {
          return mappedLanguage;
        }
      }
    }

    return null;
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

  /**
   * Map various language codes to our TargetLanguage enum
   */
  private mapLanguageCode(languageCode: string): TargetLanguage {
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
      1044: TargetLanguage.NorwegianNO,
    };

    return lcidMap[lcid] || TargetLanguage.EnglishUS;
  }

  /**
   * Filter and prioritize alerts based on user's preferred language with fallback logic
   */
  public filterAlertsForUser(
    alerts: IAlertItem[],
    userLanguage: TargetLanguage,
    policy?: ILanguagePolicy,
  ): IAlertItem[] {
    const tenantDefault = this.getTenantDefaultLanguage();
    const effectivePolicy = normalizeLanguagePolicy(policy);

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
        // Handle standalone alerts (non-multi-language)
        // Show if: targetLanguage is "all" (case-insensitive), matches user's language, or matches tenant default
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

      // Try to find alert in user's preferred language
      let selectedAlert = candidateAlerts.find(
        (alert) => alert.targetLanguage === userLanguage,
      );

      // If not found, try to find alert marked as "available for all"
      if (!selectedAlert) {
        const availableForAllAlert = candidateAlerts.find(
          (alert) => alert.availableForAll,
        );
        if (availableForAllAlert) {
          selectedAlert = availableForAllAlert;
        }
      }

      // If still not found, fall back to configured policy language
      if (!selectedAlert) {
        const fallbackLanguage =
          effectivePolicy.fallbackLanguage === "tenant-default"
            ? tenantDefault
            : effectivePolicy.fallbackLanguage;
        selectedAlert = candidateAlerts.find(
          (alert) => alert.targetLanguage === fallbackLanguage,
        );
      }

      // If still not found, fall back to tenant default language
      if (!selectedAlert) {
        selectedAlert = candidateAlerts.find(
          (alert) => alert.targetLanguage === tenantDefault,
        );
      }

      // Last resort: pick the first available alert in the group
      if (!selectedAlert) {
        selectedAlert = candidateAlerts[0];
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

  /**
   * Create a multi-language alert with content for each language
   */
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

  /**
   * Validate multi-language content has at least one complete language
   */
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

  /**
   * Generate individual alert items from multi-language alert
   */
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

  /**
   * Get language-specific content for editing multi-language alerts
   * Deduplicates by language to ensure each language appears only once
   */
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

  public detectDuplicateLanguages(
    alerts: IAlertItem[],
    languageGroup: string,
  ): TargetLanguage[] {
    const groupAlerts = alerts.filter(
      (alert) => alert.languageGroup === languageGroup,
    );
    const seen = new Set<string>();
    const duplicates = new Set<TargetLanguage>();
    groupAlerts.forEach((alert) => {
      const lang = alert.targetLanguage;
      if (seen.has(lang)) {
        duplicates.add(lang);
      } else {
        seen.add(lang);
      }
    });
    return Array.from(duplicates);
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
