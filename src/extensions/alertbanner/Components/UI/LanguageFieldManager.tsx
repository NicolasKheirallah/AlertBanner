import * as React from "react";
import {
  Checkbox as FluentCheckbox,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar as FluentMessageBar,
  MessageBarType,
} from "@fluentui/react";
import { logger } from '../Services/LoggerService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import {
  Globe24Regular,
  Add24Regular,
  Checkmark24Filled
} from "@fluentui/react-icons";
import { SharePointAlertService } from "../Services/SharePointAlertService";
import { LanguageAwarenessService } from "../Services/LanguageAwarenessService";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';
import styles from "./LanguageFieldManager.module.scss";

const cx = (...classes: Array<string | undefined | false>): string =>
  classes.filter(Boolean).join(" ");

const Card: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div className={styles.f2Card}>
    {children}
  </div>
);

const CardHeader: React.FC<{
  header?: React.ReactNode;
  description?: React.ReactNode;
}> = ({ header, description }) => (
  <div className={styles.f2CardHeader}>
    {header}
    {description}
  </div>
);

const CardPreview: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div>{children}</div>
);

const Text: React.FC<{
  children?: React.ReactNode;
  size?: number;
  weight?: "regular" | "medium" | "semibold" | "bold";
  className?: string;
}> = ({ children, size, weight, className }) => (
  <span
    className={cx(
      styles.f2Text,
      size === 100 && styles.f2Text100,
      size === 200 && styles.f2Text200,
      size === 300 && styles.f2Text300,
      size === 400 && styles.f2Text400,
      size === 500 && styles.f2Text500,
      weight === "regular" && styles.f2TextRegular,
      weight === "medium" && styles.f2TextMedium,
      weight === "semibold" && styles.f2TextSemibold,
      weight === "bold" && styles.f2TextBold,
      className,
    )}
  >
    {children}
  </span>
);

const Button: React.FC<{
  children?: React.ReactNode;
  icon?: React.ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  className?: string;
  appearance?: string;
}> = ({ children, icon, onClick, disabled, className }) => (
  <DefaultButton
    onRenderIcon={icon ? () => <>{icon}</> : undefined}
    onClick={onClick}
    disabled={disabled}
    className={cx(styles.f2Button, className)}
  >
    {children}
  </DefaultButton>
);

const Checkbox: React.FC<{
  checked?: boolean;
  disabled?: boolean;
  onChange?: (
    event: React.FormEvent<HTMLElement> | undefined,
    data: { checked?: boolean },
  ) => void;
}> = ({ checked, disabled, onChange }) => (
  <FluentCheckbox
    checked={checked}
    disabled={disabled}
    onChange={(event, isChecked) =>
      onChange?.(event as React.FormEvent<HTMLElement>, { checked: isChecked })
    }
  />
);

const Badge: React.FC<{
  children?: React.ReactNode;
  className?: string;
  icon?: React.ReactNode;
  appearance?: string;
  color?: string;
  size?: string;
}> = ({ children, className, icon }) => (
  <span className={cx(styles.f2Badge, className)}>
    {icon}
    {children}
  </span>
);

const MessageBar: React.FC<{
  intent?: "error" | "warning" | "success" | "info";
  children?: React.ReactNode;
}> = ({ intent = "info", children }) => (
  <FluentMessageBar
    messageBarType={
      intent === "error"
        ? MessageBarType.error
        : intent === "warning"
          ? MessageBarType.warning
          : intent === "success"
            ? MessageBarType.success
            : MessageBarType.info
    }
    isMultiline
  >
    {children}
  </FluentMessageBar>
);

const MessageBarBody: React.FC<{ children?: React.ReactNode }> = ({ children }) => <>{children}</>;

interface ILanguage {
  code: string;
  name: string;
  nativeName: string;
  flag: string;
  isAdded: boolean;
  isPending?: boolean;
}

interface ILanguageFieldManagerProps {
  alertService: SharePointAlertService;
  onLanguageChange?: (languages: string[]) => void;
}

const DEFAULT_LANGUAGES: ILanguage[] = [
  { code: "en-us", name: "English", nativeName: "English", flag: "🇺🇸", isAdded: true }, // Only English preselected
  { code: "fr-fr", name: "French", nativeName: "Français", flag: "🇫🇷", isAdded: false },
  { code: "sv-se", name: "Swedish", nativeName: "Svenska", flag: "🇸🇪", isAdded: false },
  { code: "de-de", name: "German", nativeName: "Deutsch", flag: "🇩🇪", isAdded: false },
  { code: "es-es", name: "Spanish", nativeName: "Español", flag: "🇪🇸", isAdded: false },
  { code: "fi-fi", name: "Finnish", nativeName: "Suomi", flag: "🇫🇮", isAdded: false },
  { code: "da-dk", name: "Danish", nativeName: "Dansk", flag: "🇩🇰", isAdded: false },
  { code: "nb-no", name: "Norwegian", nativeName: "Norsk", flag: "🇳🇴", isAdded: false }
];

const LanguageFieldManager: React.FC<ILanguageFieldManagerProps> = ({
  alertService,
  onLanguageChange
}) => {
  const [languages, setLanguages] = React.useState<ILanguage[]>(DEFAULT_LANGUAGES);
  const [message, setMessage] = React.useState<{ type: 'success' | 'error' | 'warning'; text: string } | null>(null);
  const [isTogglingLanguage, setIsTogglingLanguage] = React.useState(false);

  const getSiteDefaultLanguage = React.useCallback((): string => {
    const spLanguage = (window as any).SPClientContext?.web?.language;
    const browserLanguage = navigator.language?.toLowerCase();
    
    const languageMap: { [key: string]: string } = {
      '1033': 'en-us', // SharePoint LCID for English
      '1036': 'fr-fr', // French
      '1031': 'de-de', // German  
      '1034': 'es-es', // Spanish
      '1053': 'sv-se', // Swedish
      '1035': 'fi-fi', // Finnish
      '1030': 'da-dk', // Danish
      '1044': 'nb-no', // Norwegian
      'en': 'en-us',
      'en-us': 'en-us',
      'fr': 'fr-fr',
      'de': 'de-de',
      'es': 'es-es', 
      'sv': 'sv-se',
      'fi': 'fi-fi',
      'da': 'da-dk',
      'nb': 'nb-no',
      'no': 'nb-no'
    };
    
    if (spLanguage && languageMap[spLanguage.toString()]) {
      return languageMap[spLanguage.toString()];
    }
    
    if (browserLanguage) {
      const shortLang = browserLanguage.split('-')[0];
      if (languageMap[browserLanguage]) return languageMap[browserLanguage];
      if (languageMap[shortLang]) return languageMap[shortLang];
    }
    
    return 'en-us';
  }, []);

  const showMessage = (type: 'success' | 'error' | 'warning', text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 5000);
  };

  const { loading, execute: loadSupportedLanguages } = useAsyncOperation(
    async () => {
      const supported = await alertService.getSupportedLanguages();
      const siteDefaultLanguage = getSiteDefaultLanguage();

      logger.debug('LanguageFieldManager', `Site default language detected: ${siteDefaultLanguage}`);
      logger.debug('LanguageFieldManager', `Supported languages from SharePoint: ${supported.join(', ')}`);

      const supportedLanguages = LanguageAwarenessService.getSupportedLanguages();

      // Map our internal language list to the standardized one and update with SharePoint status
      const updatedLanguages = supportedLanguages.map(stdLang => {
        const currentLang = languages.find(l => l.code === stdLang.code);
        return {
          code: stdLang.code,
          name: stdLang.name,
          nativeName: stdLang.nativeName,
          flag: stdLang.flag,
          isAdded: supported.includes(stdLang.code) || (stdLang.code === siteDefaultLanguage && supported.length === 0),
          isPending: currentLang?.isPending || false
        };
      });

      // If no languages are supported yet and this is first load, ensure site default is selected
      if (supported.length === 0) {
        const defaultLanguages = updatedLanguages.map(lang => ({
          ...lang,
          isAdded: lang.code === siteDefaultLanguage
        }));
        logger.debug('LanguageFieldManager', `Set ${siteDefaultLanguage} as default language for new installation`);
        return defaultLanguages;
      }

      return updatedLanguages;
    },
    {
      onSuccess: (updatedLanguages) => {
        if (updatedLanguages) {
          setLanguages(updatedLanguages);
        }
      },
      onError: () => {
        logger.warn('LanguageFieldManager', 'Could not load supported languages');

        const siteDefaultLanguage = getSiteDefaultLanguage();
        setLanguages(prev => prev.map(lang => ({
          ...lang,
          isAdded: lang.code === siteDefaultLanguage
        })));

        showMessage('warning', CoreText.format(strings.LanguageManagerLoadFailedDefault, siteDefaultLanguage));
      },
      logErrors: true
    }
  );

  React.useEffect(() => {
    loadSupportedLanguages();
  }, []);

  const handleLanguageToggle = async (languageCode: string, checked: boolean) => {
    if (languageCode !== 'en-us') {
      const siteDefaultLanguage = getSiteDefaultLanguage();
      if (languageCode === siteDefaultLanguage && !checked) {
        const defaultLanguageName = languages.find(l => l.code === siteDefaultLanguage)?.name || 'default';
        showMessage('warning', CoreText.format(strings.LanguageManagerDefaultLanguageProtected, defaultLanguageName));
        return;
      }
    }

    const language = languages.find(l => l.code === languageCode);
    if (!language) return;

    setLanguages(prev => prev.map(lang => 
      lang.code === languageCode 
        ? { ...lang, isPending: true, isAdded: checked }
        : lang
    ));

    try {
      if (checked) {
        setIsTogglingLanguage(true);
        await alertService.addLanguageSupport(languageCode);
        showMessage('success', CoreText.format(strings.LanguageManagerAddedSuccess, language.name));
      } else {
        setIsTogglingLanguage(true);
        await alertService.removeLanguageSupport(languageCode);
        showMessage('success', CoreText.format(strings.LanguageManagerRemovedSuccess, language.name));
      }

      setLanguages(prev => prev.map(lang =>
        lang.code === languageCode
          ? { ...lang, isPending: false, isAdded: checked }
          : lang
      ));

      const activeLanguages = languages
        .filter(l => (l.code === languageCode ? checked : l.isAdded) && !l.isPending)
        .map(l => l.code);
      onLanguageChange?.(activeLanguages);

    } catch (error) {
      logger.error('LanguageFieldManager', `Failed to ${checked ? 'add' : 'remove'} language ${languageCode}`, error);
      showMessage('error', CoreText.format(strings.LanguageManagerUpdateFailed, language.name));

      setLanguages(prev => prev.map(lang =>
        lang.code === languageCode
          ? { ...lang, isPending: false, isAdded: !checked }
          : lang
      ));
    } finally {
      setIsTogglingLanguage(false);
    }
  };

  const getStatusBadge = (language: ILanguage) => {
    if (language.isPending) {
      return (
        <Badge 
          appearance="outline" 
          className={`${styles.statusBadge} ${styles.pendingBadge}`}
          icon={<Spinner size={SpinnerSize.xSmall} />}
        >
          {strings.LanguageManagerUpdating}
        </Badge>
      );
    }
    
    if (language.isAdded) {
      return (
        <Badge 
          appearance="filled" 
          className={`${styles.statusBadge} ${styles.addedBadge}`}
          icon={<Checkmark24Filled fontSize={12} />}
        >
          {strings.LanguageManagerActive}
        </Badge>
      );
    }
    
    return null;
  };

  const addedCount = languages.filter(l => l.isAdded && !l.isPending).length;
  const pendingCount = languages.filter(l => l.isPending).length;
  const activeLabel = pendingCount > 0
    ? CoreText.format(strings.LanguageManagerActiveCountWithPending, addedCount, pendingCount)
    : CoreText.format(strings.LanguageManagerActiveCount, addedCount);

  return (
    <div className={styles.container}>
      {message && (
        <MessageBar intent={message.type}>
          <MessageBarBody>{message.text}</MessageBarBody>
        </MessageBar>
      )}

      <Card>
        <CardHeader
          header={
            <div className={styles.f2HeaderRow}>
              <Globe24Regular />
              <Text weight="semibold">{strings.LanguageManagerTitle}</Text>
            </div>
          }
          description={
            <Text size={200}>
              {strings.LanguageManagerDescription}
            </Text>
          }
        />
        
        <CardPreview>
          <div className={styles.f2CardBody}>
            <div className={styles.languageHeader}>
              <div className={styles.languageSummary}>
                <Text weight="semibold">{strings.LanguageManagerAvailableLanguages}</Text>
                <Text size={200} className={styles.languageCount}>
                  {activeLabel}
                </Text>
              </div>
              <Button
                appearance="secondary"
                icon={<Add24Regular />}
                onClick={loadSupportedLanguages}
                disabled={loading}
                className={styles.refreshButton}
              >
                {loading ? strings.LanguageManagerLoading : strings.Refresh}
              </Button>
            </div>

            {loading ? (
              <div className={styles.loadingContainer}>
                <Spinner label={strings.LanguageManagerLoadingSupport} />
              </div>
            ) : (
              <div className={styles.languageGrid}>
                {languages.map(language => (
                  <div key={language.code} className={styles.languageItem}>
                    <div className={styles.languageInfo}>
                      <Checkbox
                        checked={language.isAdded}
                        disabled={language.isPending}
                        onChange={(_, data) => handleLanguageToggle(language.code, data.checked === true)}
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
                ))}
              </div>
            )}
          </div>
        </CardPreview>
      </Card>
    </div>
  );
};

export default LanguageFieldManager;
