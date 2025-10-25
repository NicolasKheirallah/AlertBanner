import { useMemo } from 'react';
import { TargetLanguage } from '../Alerts/IAlerts';
import { ISupportedLanguage } from '../Services/LanguageAwarenessService';

export interface ISharePointSelectOption {
  value: string;
  label: string;
}

/**
 * Custom hook to generate language options for dropdown/select components
 * @param supportedLanguages - Array of supported languages with availability status
 * @returns Array of formatted language options
 */
export const useLanguageOptions = (
  supportedLanguages: ISupportedLanguage[]
): ISharePointSelectOption[] => {
  return useMemo(() => {
    const options: ISharePointSelectOption[] = [
      { value: TargetLanguage.All, label: '🌐 All Languages' }
    ];

    // Filter to only show enabled languages (those with column support or English default)
    const enabledLanguages = supportedLanguages.filter(lang =>
      (lang.isSupported && lang.columnExists) ||
      lang.code === TargetLanguage.EnglishUS
    );

    // Add each enabled language with flag, native name, and English name
    enabledLanguages.forEach(lang => {
      options.push({
        value: lang.code,
        label: `${lang.flag} ${lang.nativeName} (${lang.name})`
      });
    });

    return options;
  }, [supportedLanguages]);
};
