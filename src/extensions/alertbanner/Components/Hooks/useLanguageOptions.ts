import { useMemo } from 'react';
import { TargetLanguage } from '../Alerts/IAlerts';
import { ISupportedLanguage } from '../Services/LanguageAwarenessService';
import { ISharePointSelectOption } from '../UI/SharePointControls';

import * as strings from 'AlertBannerApplicationCustomizerStrings';

export const useLanguageOptions = (
  supportedLanguages: ISupportedLanguage[]
): ISharePointSelectOption[] => {
  return useMemo(() => {
    const options: ISharePointSelectOption[] = [
      { value: TargetLanguage.All, label: `🌐 ${strings.CreateAlertTargetLanguageAll}` }
    ];

    const enabledLanguages = supportedLanguages.filter(lang =>
      (lang.isSupported && lang.columnExists) ||
      lang.code === TargetLanguage.EnglishUS
    );

    enabledLanguages.forEach(lang => {
      options.push({
        value: lang.code,
        label: `${lang.flag} ${lang.nativeName} (${lang.name})`
      });
    });

    return options;
  }, [supportedLanguages]);
};
