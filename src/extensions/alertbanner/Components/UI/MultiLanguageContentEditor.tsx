import * as React from "react";
import {
  Tab,
  TabList,
  Card,
  CardHeader,
  Text,
  Button,
  Field,
  Badge,
  Checkbox,
  MessageBar,
  MessageBarBody,
  Spinner,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from "@fluentui/react-components";
import {
  Add24Regular,
  Dismiss24Regular,
  Globe24Regular,
} from "@fluentui/react-icons";
import {
  SharePointInput,
  SharePointSelect,
  ISharePointSelectOption,
} from "./SharePointControls";
import SharePointRichTextEditor from "./SharePointRichTextEditor";
import {
  ILanguageContent,
  ISupportedLanguage,
} from "../Services/LanguageAwarenessService";
import { TargetLanguage, TranslationStatus } from "../Alerts/IAlerts";
import styles from "./MultiLanguageContentEditor.module.scss";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text as CoreText } from "@microsoft/sp-core-library";
import {
  ILanguagePolicy,
  normalizeLanguagePolicy,
} from "../Services/LanguagePolicyService";
import { CopilotService } from "../Services/CopilotService";
import { SparkleRegular } from "@fluentui/react-icons";
import { logger } from "../Services/LoggerService";

export interface IMultiLanguageContentEditorProps {
  content: ILanguageContent[];
  onContentChange: (content: ILanguageContent[]) => void;
  availableLanguages: ISupportedLanguage[];
  errors?: { [key: string]: string | undefined };
  linkUrl?: string;
  tenantDefaultLanguage?: TargetLanguage;
  context?: any;
  imageFolderName?: string;
  disableImageUpload?: boolean;
  languagePolicy?: ILanguagePolicy;
  copilotService?: CopilotService;
}

const MultiLanguageContentEditor: React.FC<
  IMultiLanguageContentEditorProps
> = ({
  content,
  onContentChange,
  availableLanguages,
  errors = {},
  linkUrl,
  tenantDefaultLanguage = TargetLanguage.EnglishUS,
  context,
  imageFolderName,
  disableImageUpload = false,
  languagePolicy,
  copilotService,
}) => {
  const [selectedTab, setSelectedTab] = React.useState<string>("");
  const [translatingLanguages, setTranslatingLanguages] = React.useState<
    string[]
  >([]);
  const [confirmOverwriteLang, setConfirmOverwriteLang] = React.useState<
    string | null
  >(null);
  const [translationError, setTranslationError] = React.useState<string | null>(
    null,
  );
  const [translationInfo, setTranslationInfo] = React.useState<string | null>(
    null,
  );
  const [isTranslatingAll, setIsTranslatingAll] = React.useState(false);
  const [noDefaultContentError, setNoDefaultContentError] =
    React.useState(false);
  const effectivePolicy = React.useMemo(
    () => normalizeLanguagePolicy(languagePolicy),
    [languagePolicy],
  );
  const fallbackLanguage =
    effectivePolicy.fallbackLanguage === "tenant-default"
      ? tenantDefaultLanguage
      : effectivePolicy.fallbackLanguage;

  React.useEffect(() => {
    if (content.length > 0 && !selectedTab) {
      setSelectedTab(content[0].language);
    }
  }, [content.length, selectedTab]);

  const addLanguage = (language: TargetLanguage) => {
    const languageInfo = availableLanguages.find((l) => l.code === language);
    if (!languageInfo) return;

    const newContent: ILanguageContent = {
      language,
      title: "",
      description: "",
      linkDescription: linkUrl ? "" : undefined,
      availableForAll: language === tenantDefaultLanguage,
      translationStatus: effectivePolicy.workflow.enabled
        ? effectivePolicy.workflow.defaultStatus
        : TranslationStatus.Approved,
    };

    const updatedContent = [...content, newContent];
    onContentChange(updatedContent);
    setSelectedTab(language);
  };

  const removeLanguage = (language: TargetLanguage) => {
    const updatedContent = content.filter((c) => c.language !== language);
    onContentChange(updatedContent);

    // Switch tab if we removed the current tab
    if (selectedTab === language && updatedContent.length > 0) {
      setSelectedTab(updatedContent[0].language);
    } else if (updatedContent.length === 0) {
      setSelectedTab("");
    }
  };

  const updateContent = (
    language: TargetLanguage,
    field: keyof ILanguageContent,
    value: string | boolean,
  ) => {
    const updatedContent = content.map((c) =>
      c.language === language ? { ...c, [field]: value } : c,
    );
    onContentChange(updatedContent);
  };

  const getLanguageInfo = (language: TargetLanguage) => {
    return availableLanguages.find((l) => l.code === language);
  };

  /**
   * Translates the default language content into the specified target language
   * using the CopilotService. Translates title and description in parallel.
   */
  const handleTranslate = React.useCallback(
    async (
      targetLanguage: string,
      targetLangName: string,
      overwriteExisting: boolean = true,
    ): Promise<void> => {
      if (!copilotService) return;

      const defaultContent = content.find(
        (c) => c.language === tenantDefaultLanguage,
      );
      if (
        !defaultContent ||
        (!defaultContent.title && !defaultContent.description)
      ) {
        setNoDefaultContentError(true);
        return;
      }

      setTranslatingLanguages((prev) => [...prev, targetLanguage]);

      try {
        const promises: Promise<{
          field: "title" | "description";
          value: string;
        }>[] = [];

        if (defaultContent.title) {
          promises.push(
            copilotService
              .translateText(defaultContent.title, targetLangName)
              .then((res) => ({
                field: "title" as const,
                value: res.isError ? "" : res.content,
              })),
          );
        }

        if (defaultContent.description) {
          promises.push(
            copilotService
              .translateText(defaultContent.description, targetLangName)
              .then((res) => ({
                field: "description" as const,
                value: res.isError ? "" : res.content,
              })),
          );
        }

        const results = await Promise.all(promises);

        const updatedContentList = content.map((c) => {
          if (c.language === targetLanguage) {
            const updates: Partial<ILanguageContent> = {};
            results.forEach((r) => {
              const currentValue = (c as any)[r.field] as string | undefined;
              const canWrite =
                overwriteExisting ||
                !currentValue ||
                currentValue.trim().length === 0;
              if (r.value && canWrite) {
                updates[r.field] = r.value;
              }
            });
            return { ...c, ...updates };
          }
          return c;
        });

        onContentChange(updatedContentList);
      } catch (e) {
        logger.error("MultiLanguageContentEditor", "Translation failed", e);
        setTranslationError(strings.CopilotTranslationFailed);
      } finally {
        setTranslatingLanguages((prev) =>
          prev.filter((l) => l !== targetLanguage),
        );
      }
    },
    [content, copilotService, onContentChange, tenantDefaultLanguage],
  );

  const handleTranslateAllMissing = React.useCallback(async (): Promise<void> => {
    if (!copilotService) {
      return;
    }

    setTranslationInfo(null);
    const missing = content.filter(
      (item) =>
        item.language !== tenantDefaultLanguage &&
        (!item.title.trim() || !item.description.trim()),
    );

    if (missing.length === 0) {
      setTranslationInfo(strings.MultiLanguageEditorNoMissingTranslations);
      return;
    }

    setIsTranslatingAll(true);
    try {
      await Promise.all(
        missing.map(async (item) => {
          const language = getLanguageInfo(item.language);
          await handleTranslate(
            item.language,
            language?.nativeName || item.language,
            false,
          );
        }),
      );
    } finally {
      setIsTranslatingAll(false);
    }
  }, [content, copilotService, tenantDefaultLanguage, handleTranslate]);

  const getAvailableLanguagesToAdd = () => {
    const usedLanguages = content.map((c) => c.language);
    return availableLanguages.filter(
      (lang) =>
        lang.isSupported &&
        lang.columnExists &&
        !usedLanguages.includes(lang.code),
    );
  };

  const translationStatusOptions: ISharePointSelectOption[] = React.useMemo(
    () => [
      { value: TranslationStatus.Draft, label: strings.TranslationStatusDraft },
      {
        value: TranslationStatus.InReview,
        label: strings.TranslationStatusInReview,
      },
      {
        value: TranslationStatus.Approved,
        label: strings.TranslationStatusApproved,
      },
    ],
    [],
  );

  const fallbackLanguageLabel = React.useMemo(() => {
    if (effectivePolicy.fallbackLanguage === "tenant-default") {
      return strings.LanguagePolicyFallbackTenantDefault;
    }
    const info = getLanguageInfo(fallbackLanguage);
    return info ? `${info.flag} ${info.nativeName}` : fallbackLanguage;
  }, [effectivePolicy.fallbackLanguage, fallbackLanguage, availableLanguages]);

  return (
    <div className={styles.container}>
      <Card>
        <CardHeader
          header={
            <div className={styles.cardHeader}>
              <Globe24Regular />
              <Text size={400} weight="semibold">
                {strings.MultiLanguageContent}
              </Text>
              <Badge size="small" color="informative">
                {CoreText.format(
                  strings.MultiLanguageEditorLanguageCount,
                  content.length.toString(),
                )}
              </Badge>
            </div>
          }
        />

        {/* Language Selector */}
        <div className={styles.languageSelector}>
          <Text size={300} weight="semibold">
            {strings.MultiLanguageEditorAddLanguagesLabel}
          </Text>
          {effectivePolicy.inheritance.enabled && (
            <Text size={200} className={styles.policyHint}>
              {CoreText.format(
                strings.LanguagePolicyInheritanceHint,
                fallbackLanguageLabel,
              )}
            </Text>
          )}
          <div className={styles.availableLanguages}>
            {getAvailableLanguagesToAdd().map((language) => (
              <button
                key={language.code}
                className={styles.languageButton}
                onClick={() => addLanguage(language.code)}
                type="button"
              >
                <span>{language.flag}</span>
                <span>{language.nativeName}</span>
                <Add24Regular style={{ width: "14px", height: "14px" }} />
              </button>
            ))}
          </div>
          {copilotService && content.length > 1 && (
            <div className={styles.translationActions}>
              <Button
                appearance="secondary"
                size="small"
                icon={isTranslatingAll ? <Spinner size="tiny" /> : <SparkleRegular />}
                onClick={() => void handleTranslateAllMissing()}
                disabled={isTranslatingAll}
              >
                {isTranslatingAll
                  ? strings.CopilotTranslatingLabel
                  : strings.MultiLanguageEditorTranslateAllMissing}
              </Button>
            </div>
          )}
          {getAvailableLanguagesToAdd().length === 0 && (
            <Text size={200} className={styles.allLanguagesText}>
              {strings.MultiLanguageEditorAllLanguagesAdded}
            </Text>
          )}
          {translationInfo && (
            <MessageBar intent="info">
              <MessageBarBody>{translationInfo}</MessageBarBody>
            </MessageBar>
          )}
          {(errors.languageDuplicate || errors.languageContent) && (
            <MessageBar intent="error">
              <MessageBarBody>
                {errors.languageDuplicate || errors.languageContent}
              </MessageBarBody>
            </MessageBar>
          )}
        </div>

        {/* Content Tabs */}
        {content.length > 0 ? (
          <div className={styles.tabsContainer}>
            <TabList
              selectedValue={selectedTab}
              onTabSelect={(_, data) => setSelectedTab(data.value as string)}
            >
              {content.map((contentItem) => {
                const langInfo = getLanguageInfo(contentItem.language);
                return (
                  <Tab key={contentItem.language} value={contentItem.language}>
                    <span className={styles.tabFlag}>{langInfo?.flag}</span>
                    {langInfo?.nativeName}
                    {(!contentItem.title || !contentItem.description) && (
                      <Badge
                        size="small"
                        color="warning"
                        className={styles.tabBadge}
                      >
                        {strings.MultiLanguageEditorIncompleteBadge}
                      </Badge>
                    )}
                  </Tab>
                );
              })}
            </TabList>

            {/* Tab Content */}
            {content.map((contentItem) => {
              if (selectedTab !== contentItem.language) return null;

              const langInfo = getLanguageInfo(contentItem.language);
              return (
                <div key={contentItem.language} className={styles.tabContent}>
                  <div className={styles.tabHeader}>
                    <div className={styles.languageInfo}>
                      <span>{langInfo?.flag}</span>
                      <span>{langInfo?.nativeName}</span>
                      <span className={styles.langCode}>
                        ({langInfo?.name})
                      </span>
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

                  {copilotService &&
                    contentItem.language !== tenantDefaultLanguage && (
                      <div style={{ marginBottom: "16px" }}>
                        <Button
                          appearance="subtle"
                          disabled={translatingLanguages.includes(
                            contentItem.language,
                          )}
                          icon={
                            translatingLanguages.includes(
                              contentItem.language,
                            ) ? (
                              <Spinner size="tiny" />
                            ) : (
                              <SparkleRegular />
                            )
                          }
                          size="small"
                          title={strings.CopilotDraftButton}
                          onClick={() => {
                            // Check if content exists that would be overwritten
                            if (contentItem.title || contentItem.description) {
                              setConfirmOverwriteLang(contentItem.language);
                              return;
                            }

                            // No existing content, translate directly
                            handleTranslate(
                              contentItem.language,
                              langInfo?.nativeName || contentItem.language,
                            );
                          }}
                        >
                          {translatingLanguages.includes(contentItem.language)
                            ? strings.CopilotTranslatingLabel
                            : strings.CopilotTranslateButton}
                        </Button>
                      </div>
                    )}

                  <div className={styles.contentFields}>
                    <Field
                      label={strings.MultiLanguageEditorTitleLabel}
                      required
                      validationState={
                        errors[`title_${contentItem.language}`]
                          ? "error"
                          : undefined
                      }
                      validationMessage={
                        errors[`title_${contentItem.language}`]
                      }
                    >
                      {/** Determine placeholder text using available language info */}
                      <SharePointInput
                        label=""
                        placeholder={CoreText.format(
                          strings.MultiLanguageEditorTitlePlaceholder,
                          langInfo?.nativeName ||
                            langInfo?.name ||
                            contentItem.language,
                        )}
                        value={contentItem.title}
                        onChange={(value) =>
                          updateContent(contentItem.language, "title", value)
                        }
                        error={errors[`title_${contentItem.language}`]}
                      />
                    </Field>

                    <Field
                      label={strings.MultiLanguageEditorDescriptionLabel}
                      required
                      validationState={
                        errors[`description_${contentItem.language}`]
                          ? "error"
                          : undefined
                      }
                      validationMessage={
                        errors[`description_${contentItem.language}`]
                      }
                    >
                      <SharePointRichTextEditor
                        label=""
                        value={contentItem.description}
                        onChange={(value) =>
                          updateContent(
                            contentItem.language,
                            "description",
                            value,
                          )
                        }
                        placeholder={CoreText.format(
                          strings.MultiLanguageEditorDescriptionPlaceholder,
                          langInfo?.nativeName ||
                            langInfo?.name ||
                            contentItem.language,
                        )}
                        context={context}
                        imageFolderName={imageFolderName}
                        disableImageUpload={disableImageUpload}
                      />
                    </Field>

                    {linkUrl && (
                      <Field
                        label={strings.MultiLanguageEditorLinkDescriptionLabel}
                        validationState={
                          errors[`linkDescription_${contentItem.language}`]
                            ? "error"
                            : undefined
                        }
                        validationMessage={
                          errors[`linkDescription_${contentItem.language}`]
                        }
                      >
                        <SharePointInput
                          label=""
                          placeholder={CoreText.format(
                            strings.MultiLanguageEditorLinkDescriptionPlaceholder,
                            langInfo?.nativeName ||
                              langInfo?.name ||
                              contentItem.language,
                          )}
                          value={contentItem.linkDescription || ""}
                          onChange={(value) =>
                            updateContent(
                              contentItem.language,
                              "linkDescription",
                              value,
                            )
                          }
                          error={
                            errors[`linkDescription_${contentItem.language}`]
                          }
                        />
                      </Field>
                    )}

                    {effectivePolicy.workflow.enabled && (
                      <SharePointSelect
                        label={strings.TranslationStatusLabel}
                        value={
                          contentItem.translationStatus ||
                          effectivePolicy.workflow.defaultStatus
                        }
                        onChange={(value) =>
                          updateContent(
                            contentItem.language,
                            "translationStatus",
                            value,
                          )
                        }
                        options={translationStatusOptions}
                        description={strings.TranslationStatusDescription}
                      />
                    )}

                    <div className={styles.fieldRow}>
                      <Checkbox
                        checked={!!contentItem.availableForAll}
                        onChange={(_, data) =>
                          updateContent(
                            contentItem.language,
                            "availableForAll",
                            !!data.checked,
                          )
                        }
                        label={strings.MultiLanguageEditorFallbackLabel}
                      />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          <div className={styles.emptyState}>
            <Globe24Regular className={styles.emptyIcon} />
            <Text size={400} weight="semibold">
              {strings.MultiLanguageEditorNoLanguagesTitle}
            </Text>
            <Text size={300}>
              {strings.MultiLanguageEditorNoLanguagesDescription}
            </Text>
          </div>
        )}

        {/* Summary */}
        {content.length > 0 && (
          <div className={styles.summary}>
            <Text size={300} weight="semibold">
              {strings.MultiLanguageEditorSummaryTitle}
            </Text>
            <ul className={styles.summaryList}>
              {content.map((contentItem) => {
                const langInfo = getLanguageInfo(contentItem.language);
                const isComplete = contentItem.title && contentItem.description;
                return (
                  <li key={contentItem.language} className={styles.summaryItem}>
                    <span>
                      {langInfo?.flag} {langInfo?.nativeName}:{" "}
                    </span>
                    <span
                      className={
                        isComplete
                          ? styles.statusComplete
                          : styles.statusIncomplete
                      }
                    >
                      {isComplete
                        ? strings.MultiLanguageEditorSummaryComplete
                        : strings.MultiLanguageEditorSummaryIncomplete}
                    </span>
                  </li>
                );
              })}
            </ul>
          </div>
        )}
      </Card>

      {/* Overwrite Confirmation Dialog */}
      <Dialog
        open={!!confirmOverwriteLang}
        onOpenChange={(_, data) => {
          if (!data.open) setConfirmOverwriteLang(null);
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {strings.CopilotOverwriteConfirmationTitle}
            </DialogTitle>
            <DialogContent>
              {strings.CopilotOverwriteConfirmation}
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setConfirmOverwriteLang(null)}
              >
                {strings.Cancel}
              </Button>
              <Button
                appearance="primary"
                onClick={() => {
                  if (confirmOverwriteLang) {
                    const langInfo = getLanguageInfo(
                      confirmOverwriteLang as TargetLanguage,
                    );
                    handleTranslate(
                      confirmOverwriteLang,
                      langInfo?.nativeName || confirmOverwriteLang,
                    );
                  }
                  setConfirmOverwriteLang(null);
                }}
              >
                {strings.CopilotOverwriteConfirmButton}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* No Default Content Error Dialog */}
      <Dialog
        open={noDefaultContentError}
        onOpenChange={(_, data) => {
          if (!data.open) setNoDefaultContentError(false);
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.CopilotDefaultLanguageRequired}</DialogTitle>
            <DialogContent>
              {strings.CopilotDefaultLanguageRequired}
            </DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={() => setNoDefaultContentError(false)}
              >
                {strings.Close}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Translation Error Dialog */}
      <Dialog
        open={!!translationError}
        onOpenChange={(_, data) => {
          if (!data.open) setTranslationError(null);
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.CopilotTranslationFailed}</DialogTitle>
            <DialogContent>{translationError}</DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={() => setTranslationError(null)}
              >
                {strings.Close}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default MultiLanguageContentEditor;
