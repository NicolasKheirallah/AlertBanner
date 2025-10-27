import * as React from 'react';
import {
  Tab,
  TabList,
  Card,
  CardHeader,
  Text,
  Button,
  Field,
  Badge,
  Checkbox
} from '@fluentui/react-components';
import {
  Add24Regular,
  Dismiss24Regular,
  Globe24Regular
} from '@fluentui/react-icons';
import { SharePointInput } from './SharePointControls';
import SharePointRichTextEditor from './SharePointRichTextEditor';
import { ILanguageContent, ISupportedLanguage, LanguageAwarenessService } from '../Services/LanguageAwarenessService';
import { TargetLanguage } from '../Alerts/IAlerts';
import styles from './MultiLanguageContentEditor.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';


export interface IMultiLanguageContentEditorProps {
  content: ILanguageContent[];
  onContentChange: (content: ILanguageContent[]) => void;
  availableLanguages: ISupportedLanguage[];
  errors?: { [key: string]: string };
  linkUrl?: string;
  tenantDefaultLanguage?: TargetLanguage;
  context?: any;
  imageFolderName?: string; 
}

const MultiLanguageContentEditor: React.FC<IMultiLanguageContentEditorProps> = ({
  content,
  onContentChange,
  availableLanguages,
  errors = {},
  linkUrl,
  tenantDefaultLanguage = TargetLanguage.EnglishUS,
  context,
  imageFolderName
}) => {
  const [selectedTab, setSelectedTab] = React.useState<string>('');
  
  React.useEffect(() => {
    if (content.length > 0 && !selectedTab) {
      setSelectedTab(content[0].language);
    }
  }, [content.length, selectedTab]);

  const addLanguage = (language: TargetLanguage) => {
    const languageInfo = availableLanguages.find(l => l.code === language);
    if (!languageInfo) return;

    const newContent: ILanguageContent = {
      language,
      title: '',
      description: '',
      linkDescription: linkUrl ? '' : undefined,
      availableForAll: language === tenantDefaultLanguage
    };

    const updatedContent = [...content, newContent];
    onContentChange(updatedContent);
    setSelectedTab(language);
  };

  const removeLanguage = (language: TargetLanguage) => {
    const updatedContent = content.filter(c => c.language !== language);
    onContentChange(updatedContent);
    
    // Switch tab if we removed the current tab
    if (selectedTab === language && updatedContent.length > 0) {
      setSelectedTab(updatedContent[0].language);
    } else if (updatedContent.length === 0) {
      setSelectedTab('');
    }
  };

  const updateContent = (language: TargetLanguage, field: keyof ILanguageContent, value: string) => {
    const updatedContent = content.map(c => 
      c.language === language 
        ? { ...c, [field]: value }
        : c
    );
    onContentChange(updatedContent);
  };

  const getLanguageInfo = (language: TargetLanguage) => {
    return availableLanguages.find(l => l.code === language);
  };

  const getAvailableLanguagesToAdd = () => {
    const usedLanguages = content.map(c => c.language);
    return availableLanguages.filter(lang => 
      lang.isSupported && lang.columnExists && !usedLanguages.includes(lang.code)
    );
  };

  return (
    <div className={styles.container}>
      <Card>
        <CardHeader
          header={
            <div className={styles.cardHeader}>
              <Globe24Regular />
              <Text size={400} weight="semibold">{strings.MultiLanguageContent}</Text>
              <Badge size="small" color="informative">{CoreText.format(strings.MultiLanguageEditorLanguageCount, content.length.toString())}</Badge>
            </div>
          }
        />

        {/* Language Selector */}
        <div className={styles.languageSelector}>
          <Text size={300} weight="semibold">{strings.MultiLanguageEditorAddLanguagesLabel}</Text>
          <div className={styles.availableLanguages}>
            {getAvailableLanguagesToAdd().map(language => (
              <button
                key={language.code}
                className={styles.languageButton}
                onClick={() => addLanguage(language.code)}
                type="button"
              >
                <span>{language.flag}</span>
                <span>{language.nativeName}</span>
                <Add24Regular style={{ width: '14px', height: '14px' }} />
              </button>
            ))}
          </div>
          {getAvailableLanguagesToAdd().length === 0 && (
            <Text size={200} className={styles.allLanguagesText}>
              {strings.MultiLanguageEditorAllLanguagesAdded}
            </Text>
          )}
        </div>

        {/* Content Tabs */}
        {content.length > 0 ? (
          <div className={styles.tabsContainer}>
            <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)}>
              {content.map(contentItem => {
                const langInfo = getLanguageInfo(contentItem.language);
                return (
                  <Tab key={contentItem.language} value={contentItem.language}>
                    <span className={styles.tabFlag}>{langInfo?.flag}</span>
                    {langInfo?.nativeName}
                    {(!contentItem.title || !contentItem.description) && (
                      <Badge size="small" color="warning" className={styles.tabBadge}>{strings.MultiLanguageEditorIncompleteBadge}</Badge>
                    )}
                  </Tab>
                );
              })}
            </TabList>

            {/* Tab Content */}
            {content.map(contentItem => {
              if (selectedTab !== contentItem.language) return null;
              
              const langInfo = getLanguageInfo(contentItem.language);
              return (
                <div key={contentItem.language} className={styles.tabContent}>
                  <div className={styles.tabHeader}>
                    <div className={styles.languageInfo}>
                      <span>{langInfo?.flag}</span>
                      <span>{langInfo?.nativeName}</span>
                      <span className={styles.langCode}>({langInfo?.name})</span>
                    </div>
                      {content.length > 1 && (
                        <Button
                          appearance="subtle"
                          icon={<Dismiss24Regular />}
                          onClick={() => removeLanguage(contentItem.language)}
                          className={styles.removeButton}
                          size="small"
                        >
                          {strings.MultiLanguageEditorRemoveButton}
                        </Button>
                      )}
                    </div>

                    <div className={styles.contentFields}>
                      <Field
                        label={strings.MultiLanguageEditorTitleLabel}
                        required
                        validationState={errors[`title_${contentItem.language}`] ? 'error' : undefined}
                        validationMessage={errors[`title_${contentItem.language}`]}
                      >
                        {/** Determine placeholder text using available language info */}
                        <SharePointInput
                          label=""
                          placeholder={CoreText.format(strings.MultiLanguageEditorTitlePlaceholder, (langInfo?.nativeName || langInfo?.name || contentItem.language))}
                          value={contentItem.title}
                          onChange={(value) => updateContent(contentItem.language, 'title', value)}
                          error={errors[`title_${contentItem.language}`]}
                        />
                      </Field>

                      <Field
                        label={strings.MultiLanguageEditorDescriptionLabel}
                        required
                        validationState={errors[`description_${contentItem.language}`] ? 'error' : undefined}
                        validationMessage={errors[`description_${contentItem.language}`]}
                      >
                        <SharePointRichTextEditor
                          label=""
                          value={contentItem.description}
                          onChange={(value) => updateContent(contentItem.language, 'description', value)}
                          placeholder={CoreText.format(strings.MultiLanguageEditorDescriptionPlaceholder, (langInfo?.nativeName || langInfo?.name || contentItem.language))}
                          context={context}
                          imageFolderName={imageFolderName}
                        />
                      </Field>

                      {linkUrl && (
                        <Field
                          label={strings.MultiLanguageEditorLinkDescriptionLabel}
                          validationState={errors[`linkDescription_${contentItem.language}`] ? 'error' : undefined}
                          validationMessage={errors[`linkDescription_${contentItem.language}`]}
                        >
                              <SharePointInput
                                label=""
                                placeholder={CoreText.format(strings.MultiLanguageEditorLinkDescriptionPlaceholder, (langInfo?.nativeName || langInfo?.name || contentItem.language))}
                                value={contentItem.linkDescription || ''}
                                onChange={(value) => updateContent(contentItem.language, 'linkDescription', value)}
                                error={errors[`linkDescription_${contentItem.language}`]}
                              />
                        </Field>
                      )}
                    </div>
                  </div>
                );
            })}
          </div>
          ) : (
          <div className={styles.emptyState}>
            <Globe24Regular className={styles.emptyIcon} />
            <Text size={400} weight="semibold">{strings.MultiLanguageEditorNoLanguagesTitle}</Text>
            <Text size={300}>{strings.MultiLanguageEditorNoLanguagesDescription}</Text>
          </div>
        )}

        {/* Summary */}
        {content.length > 0 && (
          <div className={styles.summary}>
            <Text size={300} weight="semibold">{strings.MultiLanguageEditorSummaryTitle}</Text>
            <ul className={styles.summaryList}>
              {content.map(contentItem => {
                const langInfo = getLanguageInfo(contentItem.language);
                const isComplete = contentItem.title && contentItem.description;
                return (
                  <li key={contentItem.language} className={styles.summaryItem}>
                    <span>{langInfo?.flag} {langInfo?.nativeName}: </span>
                    <span className={isComplete ? styles.statusComplete : styles.statusIncomplete}>
                      {isComplete ? strings.MultiLanguageEditorSummaryComplete : strings.MultiLanguageEditorSummaryIncomplete}
                    </span>
                  </li>
                );
              })}
            </ul>
          </div>
        )}
      </Card>
    </div>
  );
};

export default MultiLanguageContentEditor;
