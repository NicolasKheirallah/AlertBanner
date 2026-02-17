import * as React from 'react';
import { logger } from '../Services/LoggerService';
import {
  Dropdown,
  DefaultButton,
  IconButton,
  IDropdownOption,
  IContextualMenuItem,
} from "@fluentui/react";
import { LocalLanguage24Regular } from '@fluentui/react-icons';
import { useLocalization } from '../Hooks/useLocalization';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import styles from './LanguageSelector.module.scss';

export interface ILanguageSelectorProps {
  compact?: boolean;
  className?: string;
  onLanguageChange?: (languageCode: string) => void;
}

const LanguageSelector: React.FC<ILanguageSelectorProps> = ({
  compact = false,
  className,
  onLanguageChange
}) => {
  const { 
    currentLanguage, 
    supportedLanguages, 
    setLanguage 
  } = useLocalization();

  const dropdownOptions = React.useMemo<IDropdownOption[]>(
    () =>
      supportedLanguages.map((language) => ({
        key: language.code,
        text: `${language.nativeName} (${language.name})`,
      })),
    [supportedLanguages],
  );

  const compactMenuItems = React.useMemo<IContextualMenuItem[]>(
    () =>
      supportedLanguages.map((language) => ({
        key: language.code,
        text: `${language.nativeName} (${language.name})`,
        disabled: language.code === currentLanguage.code,
        onClick: () => {
          void handleLanguageChange(language.code);
        },
      })),
    [supportedLanguages, currentLanguage.code],
  );

  const handleLanguageChange = async (languageCode: string) => {
    try {
      await setLanguage(languageCode);
      onLanguageChange?.(languageCode);
    } catch (error) {
      logger.error('LanguageSelector', 'Failed to change language', error);
    }
  };

  if (compact) {
    return (
      <IconButton
        onRenderIcon={() => <LocalLanguage24Regular />}
        ariaLabel={strings.ChangeLanguage}
        title={strings.ChangeLanguage}
        className={`${styles.compactButton} ${className || ""}`}
        menuProps={{ items: compactMenuItems }}
      />
    );
  }

  return (
    <div className={`${styles.languageSelector} ${className || ''}`}>
      <Dropdown
        aria-label={strings.SelectLanguage}
        placeholder={strings.SelectLanguage}
        selectedKey={currentLanguage.code}
        options={dropdownOptions}
        onChange={(_, option) => {
          if (option?.key && option.key !== currentLanguage.code) {
            void handleLanguageChange(String(option.key));
          }
        }}
      />
    </div>
  );
};

export default LanguageSelector;
