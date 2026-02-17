import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { logger } from './LoggerService';

/* eslint-disable @typescript-eslint/no-var-requires */
/* eslint-disable global-require */

export interface ILanguageInfo {
  code: string;
  name: string;
  nativeName: string;
  isRTL: boolean;
}

export interface ILocalizationStrings {
  [key: string]: string;
}

const LANGUAGE_STRINGS: { [key: string]: ILocalizationStrings } = {
  'en-us': require('../../loc/en-us.js') as ILocalizationStrings,
  'fr-fr': require('../../loc/fr-fr.js') as ILocalizationStrings,
  'de-de': require('../../loc/de-de.js') as ILocalizationStrings,
  'es-es': require('../../loc/es-es.js') as ILocalizationStrings,
  'sv-se': require('../../loc/sv-se.js') as ILocalizationStrings,
  'fi-fi': require('../../loc/fi-fi.js') as ILocalizationStrings,
  'da-dk': require('../../loc/da-dk.js') as ILocalizationStrings,
  'nb-no': require('../../loc/nb-no.js') as ILocalizationStrings
};

export class LocalizationService {
  private static _instance: LocalizationService;
  private _currentLanguage: string = 'en-us';
  private _strings: ILocalizationStrings = {};
  private _context: ApplicationCustomizerContext;
  private _fallbackStrings: ILocalizationStrings = {};

  // Supported languages
  private readonly _supportedLanguages: ILanguageInfo[] = [
    {
      code: 'en-us',
      name: 'English',
      nativeName: 'English',
      isRTL: false
    },
    {
      code: 'fr-fr',
      name: 'French',
      nativeName: 'Français',
      isRTL: false
    },
    {
      code: 'de-de',
      name: 'German',
      nativeName: 'Deutsch',
      isRTL: false
    },
    {
      code: 'es-es',
      name: 'Spanish',
      nativeName: 'Español',
      isRTL: false
    },
    {
      code: 'sv-se',
      name: 'Swedish',
      nativeName: 'Svenska',
      isRTL: false
    },
    {
      code: 'fi-fi',
      name: 'Finnish',
      nativeName: 'Suomi',
      isRTL: false
    },
    {
      code: 'da-dk',
      name: 'Danish',
      nativeName: 'Dansk',
      isRTL: false
    },
    {
      code: 'nb-no',
      name: 'Norwegian',
      nativeName: 'Norsk bokmål',
      isRTL: false
    }
  ];

  public static getInstance(context?: ApplicationCustomizerContext): LocalizationService {
    if (!LocalizationService._instance) {
      LocalizationService._instance = new LocalizationService(context);
    }
    return LocalizationService._instance;
  }

  private constructor(context?: ApplicationCustomizerContext) {
    if (context) {
      this._context = context;
    }
  }

  public async initialize(context?: ApplicationCustomizerContext): Promise<void> {
    if (context) {
      this._context = context;
    }

    const detectedLanguage = this.detectUserLanguage();
    this._currentLanguage = this.getSupportedLanguage(detectedLanguage);

    await this.loadLanguageStrings(this._currentLanguage);
    
    if (this._currentLanguage !== 'en-us') {
      await this.loadFallbackStrings();
    }
  }

  public getString(key: string, ...args: any[]): string {
    let value = this._strings[key];
    
    if (!value && this._fallbackStrings[key]) {
      value = this._fallbackStrings[key];
    }

    if (!value) {
      logger.warn('LocalizationService', `Localization key not found: ${key}`);
      return key;
    }

    if (args && args.length > 0) {
      return this.formatString(value, ...args);
    }

    return value;
  }

  public getSupportedLanguages(): ILanguageInfo[] {
    return [...this._supportedLanguages];
  }

  public getCurrentLanguage(): ILanguageInfo {
    return this._supportedLanguages.find(lang => lang.code === this._currentLanguage) || this._supportedLanguages[0];
  }

  public async setLanguage(languageCode: string): Promise<void> {
    const supportedLanguage = this.getSupportedLanguage(languageCode);
    
    if (supportedLanguage !== this._currentLanguage) {
      this._currentLanguage = supportedLanguage;
      await this.loadLanguageStrings(this._currentLanguage);
      
      // Store user preference
      this.storeLanguagePreference(this._currentLanguage);
    }
  }

  public isRTL(): boolean {
    const currentLang = this.getCurrentLanguage();
    return currentLang ? currentLang.isRTL : false;
  }

  public formatDate(date: Date | string, options?: Intl.DateTimeFormatOptions): string {
    const dateObj = typeof date === 'string' ? new Date(date) : date;
    const locale = this._currentLanguage.replace('-', '-');
    
    const defaultOptions: Intl.DateTimeFormatOptions = {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      ...options
    };

    try {
      return dateObj.toLocaleDateString(locale, defaultOptions);
    } catch (error) {
      return dateObj.toLocaleDateString('en-US', defaultOptions);
    }
  }

  public formatTime(date: Date | string, options?: Intl.DateTimeFormatOptions): string {
    const dateObj = typeof date === 'string' ? new Date(date) : date;
    const locale = this._currentLanguage.replace('-', '-');
    
    const defaultOptions: Intl.DateTimeFormatOptions = {
      hour: '2-digit',
      minute: '2-digit',
      ...options
    };

    try {
      return dateObj.toLocaleTimeString(locale, defaultOptions);
    } catch (error) {
      return dateObj.toLocaleTimeString('en-US', defaultOptions);
    }
  }

  public formatRelativeTime(date: Date | string): string {
    const dateObj = typeof date === 'string' ? new Date(date) : date;
    const now = new Date();
    const diffInSeconds = Math.floor((now.getTime() - dateObj.getTime()) / 1000);

    if (Intl.RelativeTimeFormat) {
      try {
        const rtf = new Intl.RelativeTimeFormat(this._currentLanguage.replace('-', '-'), { numeric: 'auto' });
        
        if (diffInSeconds < 60) {
          return rtf.format(-diffInSeconds, 'second');
        } else if (diffInSeconds < 3600) {
          return rtf.format(-Math.floor(diffInSeconds / 60), 'minute');
        } else if (diffInSeconds < 86400) {
          return rtf.format(-Math.floor(diffInSeconds / 3600), 'hour');
        } else {
          return rtf.format(-Math.floor(diffInSeconds / 86400), 'day');
        }
      } catch (error) {
      }
    }

    if (diffInSeconds < 60) {
      return this.getString('JustNow', diffInSeconds);
    } else if (diffInSeconds < 3600) {
      const minutes = Math.floor(diffInSeconds / 60);
      return this.getString('MinutesAgo', minutes);
    } else if (diffInSeconds < 86400) {
      const hours = Math.floor(diffInSeconds / 3600);
      return this.getString('HoursAgo', hours);
    } else {
      const days = Math.floor(diffInSeconds / 86400);
      return this.getString('DaysAgo', days);
    }
  }

  private detectUserLanguage(): string {
    let language = 'en-us';

    try {
      if (this._context?.pageContext?.cultureInfo?.currentUICultureName) {
        language = this._context.pageContext.cultureInfo.currentUICultureName.toLowerCase();
      }
      else if (navigator.language) {
        language = navigator.language.toLowerCase();
      }
      else {
        const stored = localStorage.getItem('alertbanner-language');
        if (stored) {
          language = stored;
        }
      }
    } catch (error) {
      logger.warn('LocalizationService', 'Error detecting user language', error);
    }

    return language;
  }

  private getSupportedLanguage(languageCode: string): string {
    const normalizedCode = languageCode.toLowerCase();
    
    if (this._supportedLanguages.some(lang => lang.code === normalizedCode)) {
      return normalizedCode;
    }

    const languageOnly = normalizedCode.split('-')[0];
    const matchedLanguage = this._supportedLanguages.find(lang => 
      lang.code.startsWith(languageOnly)
    );

    return matchedLanguage ? matchedLanguage.code : 'en-us';
  }

  private async loadLanguageStrings(languageCode: string): Promise<void> {
    this._strings = this.getLanguageStrings(languageCode);
    logger.debug('LocalizationService', `Loaded localization strings for ${languageCode}`);
  }

  private getLanguageStrings(languageCode: string): ILocalizationStrings {
    return LANGUAGE_STRINGS[languageCode] || LANGUAGE_STRINGS['en-us'] || {};
  }

  private async loadFallbackStrings(): Promise<void> {
    try {
      this._fallbackStrings = this.getLanguageStrings('en-us');
    } catch (error) {
      logger.warn('LocalizationService', 'Failed to load fallback language strings', error);
      this._fallbackStrings = {};
    }
  }

  private formatString(template: string, ...args: any[]): string {
    return template.replace(/\{(\d+)\}/g, (match, index) => {
      const argIndex = parseInt(index, 10);
      return args[argIndex] !== undefined ? String(args[argIndex]) : match;
    });
  }

  private storeLanguagePreference(languageCode: string): void {
    try {
      localStorage.setItem('alertbanner-language', languageCode);
    } catch (error) {
      logger.warn('LocalizationService', 'Failed to store language preference', error);
    }
  }
}

// Export a convenience function for getting localized strings
export function getString(key: string, ...args: any[]): string {
  return LocalizationService.getInstance().getString(key, ...args);
}

// Export a convenience function for formatting dates
export function formatDate(date: Date | string, options?: Intl.DateTimeFormatOptions): string {
  return LocalizationService.getInstance().formatDate(date, options);
}

// Export a convenience function for formatting relative time
export function formatRelativeTime(date: Date | string): string {
  return LocalizationService.getInstance().formatRelativeTime(date);
}
