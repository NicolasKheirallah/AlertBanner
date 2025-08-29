import * as React from "react";
import {
  Checkbox,
} from "@fluentui/react-components";

interface ILanguage {
  code: string;
  name: string;
  nativeName: string;
  flag: string;
  isAdded: boolean;
  isPending?: boolean;
}

interface ILanguageItemProps {
  language: ILanguage;
  styles: any;
  onToggle: (languageCode: string, checked: boolean) => void;
  getStatusBadge: (language: ILanguage) => React.ReactNode;
}

const LanguageItem: React.FC<ILanguageItemProps> = React.memo(({
  language,
  styles,
  onToggle,
  getStatusBadge
}) => {
  const handleToggle = React.useCallback((_, data) => {
    onToggle(language.code, data.checked === true);
  }, [language.code, onToggle]);

  return (
    <div className={styles.languageItem}>
      <div className={styles.languageInfo}>
        <Checkbox
          checked={language.isAdded}
          disabled={language.code === 'en-us' || language.isPending}
          onChange={handleToggle}
        />
        <div className={styles.languageDetails}>
          <div className={styles.languageName}>
            {language.flag} {language.nativeName}
          </div>
          <div className={styles.languageCode}>
            {language.name} ({language.code.toUpperCase()})
          </div>
        </div>
      </div>
      {getStatusBadge(language)}
    </div>
  );
});

LanguageItem.displayName = 'LanguageItem';

export default LanguageItem;