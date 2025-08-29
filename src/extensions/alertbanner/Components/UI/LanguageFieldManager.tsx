import * as React from "react";
import {
  Card,
  CardHeader,
  CardPreview,
  Text,
  Button,
  Checkbox,
  Spinner,
  MessageBar,
  MessageBarBody,
  Badge,
  makeStyles,
  tokens
} from "@fluentui/react-components";
import {
  Globe24Regular,
  Add24Regular,
  Checkmark24Filled
} from "@fluentui/react-icons";
import { SharePointAlertService } from "../Services/SharePointAlertService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "20px"
  },
  languageGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))",
    gap: "12px",
    marginTop: "16px",
    marginRight: "20px" // Prevent cutoff on right side
  },
  languageItem: {
    padding: "12px 16px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px"
  },
  languageInfo: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flex: 1
  },
  languageDetails: {
    flex: 1
  },
  languageName: {
    fontWeight: "500",
    fontSize: "14px",
    marginBottom: "2px"
  },
  languageCode: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground2
  },
  statusBadge: {
    fontSize: "10px",
    padding: "2px 6px"
  },
  addedBadge: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1
  },
  pendingBadge: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    color: tokens.colorPaletteYellowForeground1
  },
  actions: {
    display: "flex",
    gap: "8px",
    paddingTop: "16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    marginTop: "20px"
  }
});

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
  { code: "en-us", name: "English (US)", nativeName: "English", flag: "🇺🇸", isAdded: true }, // Only English preselected
  { code: "fr-fr", name: "French (France)", nativeName: "Français", flag: "🇫🇷", isAdded: false },
  { code: "sv-se", name: "Swedish (Sweden)", nativeName: "Svenska", flag: "🇸🇪", isAdded: false },
  { code: "de-de", name: "German (Germany)", nativeName: "Deutsch", flag: "🇩🇪", isAdded: false },
  { code: "es-es", name: "Spanish (Spain)", nativeName: "Español", flag: "🇪🇸", isAdded: false },
  { code: "fi-fi", name: "Finnish (Finland)", nativeName: "Suomi", flag: "🇫🇮", isAdded: false },
  { code: "da-dk", name: "Danish (Denmark)", nativeName: "Dansk", flag: "🇩🇰", isAdded: false },
  { code: "nb-no", name: "Norwegian (Norway)", nativeName: "Norsk", flag: "🇳🇴", isAdded: false }
];

const LanguageFieldManager: React.FC<ILanguageFieldManagerProps> = ({
  alertService,
  onLanguageChange
}) => {
  const styles = useStyles();
  const [languages, setLanguages] = React.useState<ILanguage[]>(DEFAULT_LANGUAGES);
  const [loading, setLoading] = React.useState(false);
  const [message, setMessage] = React.useState<{ type: 'success' | 'error' | 'warning'; text: string } | null>(null);
  
  // Get site's default language
  const getSiteDefaultLanguage = React.useCallback((): string => {
    // Try to get from SharePoint context, fallback to browser, then English
    const spLanguage = (window as any).SPClientContext?.web?.language;
    const browserLanguage = navigator.language?.toLowerCase();
    
    // Map common language codes to our supported ones
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
    
    // Try SharePoint language first
    if (spLanguage && languageMap[spLanguage.toString()]) {
      return languageMap[spLanguage.toString()];
    }
    
    // Try browser language
    if (browserLanguage) {
      const shortLang = browserLanguage.split('-')[0];
      if (languageMap[browserLanguage]) return languageMap[browserLanguage];
      if (languageMap[shortLang]) return languageMap[shortLang];
    }
    
    // Default to English
    return 'en-us';
  }, []);

  // Load currently supported languages on component mount
  React.useEffect(() => {
    loadSupportedLanguages();
  }, []);

  const showMessage = (type: 'success' | 'error' | 'warning', text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 5000);
  };

  const loadSupportedLanguages = async () => {
    try {
      setLoading(true);
      const supported = await alertService.getSupportedLanguageColumns();
      const siteDefaultLanguage = getSiteDefaultLanguage();
      
      console.log(`🌍 Site default language detected: ${siteDefaultLanguage}`);
      console.log(`📋 Supported languages from SharePoint: ${supported.join(', ')}`);
      
      // Update language states: only add languages that exist in SharePoint OR the site default
      setLanguages(prev => prev.map(lang => ({
        ...lang,
        isAdded: supported.includes(lang.code) || (lang.code === siteDefaultLanguage && supported.length === 0)
      })));
      
      // If no languages are supported yet and this is first load, ensure site default is selected
      if (supported.length === 0) {
        setLanguages(prev => prev.map(lang => ({
          ...lang,
          isAdded: lang.code === siteDefaultLanguage
        })));
        console.log(`✅ Set ${siteDefaultLanguage} as default language for new installation`);
      }
      
    } catch (error) {
      console.warn('Could not load supported languages:', error);
      
      // Fallback: set only site default language as active
      const siteDefaultLanguage = getSiteDefaultLanguage();
      setLanguages(prev => prev.map(lang => ({
        ...lang,
        isAdded: lang.code === siteDefaultLanguage
      })));
      
      showMessage('warning', `Could not load language support. Using site default: ${siteDefaultLanguage}`);
    } finally {
      setLoading(false);
    }
  };

  const handleLanguageToggle = async (languageCode: string, checked: boolean) => {
    const siteDefaultLanguage = getSiteDefaultLanguage();
    if (languageCode === siteDefaultLanguage && !checked) {
      const defaultLanguageName = languages.find(l => l.code === siteDefaultLanguage)?.name || 'default';
      showMessage('warning', `${defaultLanguageName} is the site's default language and cannot be removed.`);
      return;
    }

    const language = languages.find(l => l.code === languageCode);
    if (!language) return;

    // Update UI immediately to show pending state
    setLanguages(prev => prev.map(lang => 
      lang.code === languageCode 
        ? { ...lang, isPending: true, isAdded: checked }
        : lang
    ));

    try {
      if (checked) {
        // Add language columns
        setLoading(true);
        await alertService.addLanguageColumns(languageCode);
        showMessage('success', `Added ${language.name} language support successfully!`);
      } else {
        // Remove language columns from SharePoint
        setLoading(true);
        await alertService.removeLanguageColumns(languageCode);
        showMessage('success', `Removed ${language.name} language support and columns.`);
      }

      // Update final state
      setLanguages(prev => prev.map(lang => 
        lang.code === languageCode 
          ? { ...lang, isPending: false, isAdded: checked }
          : lang
      ));

      // Notify parent component
      const activeLanguages = languages
        .filter(l => (l.code === languageCode ? checked : l.isAdded) && !l.isPending)
        .map(l => l.code);
      onLanguageChange?.(activeLanguages);

    } catch (error) {
      console.error(`Failed to ${checked ? 'add' : 'remove'} language ${languageCode}:`, error);
      showMessage('error', `Failed to ${checked ? 'add' : 'remove'} ${language.name} language support.`);
      
      // Revert UI state on error
      setLanguages(prev => prev.map(lang => 
        lang.code === languageCode 
          ? { ...lang, isPending: false, isAdded: !checked }
          : lang
      ));
    } finally {
      setLoading(false);
    }
  };

  const getStatusBadge = (language: ILanguage) => {
    if (language.isPending) {
      return (
        <Badge 
          appearance="outline" 
          className={`${styles.statusBadge} ${styles.pendingBadge}`}
          icon={<Spinner size="tiny" />}
        >
          Updating...
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
          Active
        </Badge>
      );
    }
    
    return null;
  };

  const addedCount = languages.filter(l => l.isAdded && !l.isPending).length;
  const pendingCount = languages.filter(l => l.isPending).length;

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
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Globe24Regular />
              <Text weight="semibold">Multi-Language Field Management</Text>
            </div>
          }
          description={
            <Text size={200}>
              Select languages to add multi-language content fields to your alert lists. 
              Fields will be created for Title, Description, and Link Description in each selected language.
            </Text>
          }
        />
        
        <CardPreview>
          <div style={{ padding: "16px" }}>
            <div style={{ 
              display: 'flex', 
              alignItems: 'center', 
              justifyContent: 'space-between', 
              marginBottom: '16px',
              flexWrap: 'wrap',
              gap: '8px'
            }}>
              <div style={{ minWidth: '0', flex: '1' }}>
                <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '4px' }}>
                  <Text weight="semibold">Available Languages</Text>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground2 }}>
                    {addedCount} active{pendingCount > 0 ? `, ${pendingCount} updating` : ''}
                  </Text>
                </div>
              </div>
              <Button
                appearance="secondary"
                icon={<Add24Regular />}
                onClick={loadSupportedLanguages}
                disabled={loading}
                style={{ flexShrink: 0 }}
              >
                {loading ? 'Loading...' : 'Refresh'}
              </Button>
            </div>

            {loading ? (
              <div style={{ textAlign: 'center', padding: '40px' }}>
                <Spinner label="Loading language support..." />
              </div>
            ) : (
              <div className={styles.languageGrid}>
                {languages.map(language => (
                  <div key={language.code} className={styles.languageItem}>
                    <div className={styles.languageInfo}>
                      <Checkbox
                        checked={language.isAdded}
                        disabled={language.code === 'en-us' || language.isPending}
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

            <div className={styles.actions}>
              <Text size={200} style={{ flex: 1, color: tokens.colorNeutralForeground2 }}>
                💡 Tip: Language fields are created as: Title_{'{LANG}'}, Description_{'{LANG}'}, LinkDescription_{'{LANG}'}
              </Text>
            </div>
          </div>
        </CardPreview>
      </Card>
    </div>
  );
};

export default LanguageFieldManager;