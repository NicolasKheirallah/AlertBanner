import * as React from "react";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import {
  DefaultButton,
  PrimaryButton,
  Spinner as FluentSpinner,
  SpinnerSize,
  Checkbox as FluentCheckbox,
  MessageBar as FluentMessageBar,
  MessageBarType,
  Dialog as FluentDialog,
} from "@fluentui/react";
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
  TargetLanguage,
  TranslationStatus,
  ILanguageContent,
} from "../Alerts/IAlerts";
import { ISupportedLanguage } from "../Services/LanguageAwarenessService";
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

const cx = (...classes: Array<string | undefined | false>): string =>
  classes.filter(Boolean).join(" ");

const Card: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div className={styles.f2Card}>{children}</div>
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
  appearance?: "primary" | "secondary" | "subtle";
  size?: "small" | "medium" | "large";
  icon?: React.ReactNode;
  onClick?: () => void | Promise<void>;
  disabled?: boolean;
  title?: string;
  className?: string;
}> = ({
  children,
  appearance = "secondary",
  icon,
  onClick,
  disabled,
  title,
  className,
  size,
}) => {
  const buttonClassName = cx(
    styles.f2Button,
    appearance === "primary" && styles.f2ButtonPrimary,
    appearance === "subtle" && styles.f2ButtonSubtle,
    size === "small" && styles.f2ButtonSmall,
    className,
  );

  const commonProps = {
    onRenderIcon: icon ? () => <>{icon}</> : undefined,
    onClick,
    disabled,
    title,
    className: buttonClassName,
  };

  if (appearance === "primary") {
    return <PrimaryButton {...commonProps}>{children}</PrimaryButton>;
  }

  return <DefaultButton {...commonProps}>{children}</DefaultButton>;
};

const Field: React.FC<{
  children?: React.ReactNode;
  label?: React.ReactNode;
  required?: boolean;
  validationMessage?: React.ReactNode;
  validationState?: "error" | "warning" | "success";
}> = ({ children, label, required, validationMessage, validationState }) => (
  <div className={styles.f2Field}>
    {label ? (
      <label className={styles.f2FieldLabel}>
        {label}
        {required && <span className={styles.f2FieldRequired}> *</span>}
      </label>
    ) : null}
    {children}
    {validationMessage && (
      <div
        className={cx(
          styles.f2ValidationMessage,
          validationState === "warning" && styles.f2ValidationWarning,
          validationState === "success" && styles.f2ValidationSuccess,
          (!validationState || validationState === "error") &&
            styles.f2ValidationError,
        )}
      >
        {validationMessage}
      </div>
    )}
  </div>
);

const Badge: React.FC<{
  children?: React.ReactNode;
  className?: string;
  size?: "small" | "large";
  color?: string;
}> = ({ children, className, size }) => (
  <span
    className={cx(
      styles.f2Badge,
      size === "large" && styles.f2BadgeLarge,
      className,
    )}
  >
    {children}
  </span>
);

const Checkbox: React.FC<{
  checked?: boolean;
  label?: React.ReactNode;
  onChange?: (
    event: React.FormEvent<HTMLElement> | undefined,
    data: { checked?: boolean },
  ) => void;
}> = ({ checked, label, onChange }) => (
  <FluentCheckbox
    checked={checked}
    label={typeof label === "string" ? label : undefined}
    onRenderLabel={
      typeof label === "string" || typeof label === "undefined"
        ? undefined
        : () => <>{label}</>
    }
    onChange={(event, isChecked) =>
      onChange?.(event as React.FormEvent<HTMLElement>, { checked: isChecked })
    }
  />
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

const MessageBarBody: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <>{children}</>;

const Spinner: React.FC<{
  size?: "tiny" | "small" | "medium" | "large";
  label?: string;
}> = ({ size = "medium", label }) => (
  <FluentSpinner
    size={
      size === "tiny"
        ? SpinnerSize.xSmall
        : size === "small"
          ? SpinnerSize.small
          : size === "large"
            ? SpinnerSize.large
            : SpinnerSize.medium
    }
    label={label}
  />
);

const TabContext = React.createContext<{
  selectedValue?: string;
  onTabSelect?: (
    event: React.MouseEvent<HTMLButtonElement>,
    data: { value: string },
  ) => void;
}>({});

const TabList: React.FC<{
  selectedValue?: string;
  onTabSelect?: (
    event: React.MouseEvent<HTMLButtonElement>,
    data: { value: string },
  ) => void;
  children?: React.ReactNode;
}> = ({ selectedValue, onTabSelect, children }) => (
  <TabContext.Provider value={{ selectedValue, onTabSelect }}>
    <div role="tablist" className={styles.f2TabList}>
      {children}
    </div>
  </TabContext.Provider>
);

const Tab: React.FC<{ value: string; children?: React.ReactNode }> = ({
  value,
  children,
}) => {
  const { selectedValue, onTabSelect } = React.useContext(TabContext);
  const selected = selectedValue === value;
  return (
    <button
      type="button"
      role="tab"
      aria-selected={selected}
      onClick={(event) => onTabSelect?.(event, { value })}
      className={cx(styles.f2Tab, selected && styles.f2TabSelected)}
    >
      {children}
    </button>
  );
};

const Dialog: React.FC<{
  open: boolean;
  onOpenChange?: (
    event: React.SyntheticEvent<HTMLElement> | undefined,
    data: { open: boolean },
  ) => void;
  children?: React.ReactNode;
}> = ({ open, onOpenChange, children }) => (
  <FluentDialog
    hidden={!open}
    onDismiss={(event) =>
      onOpenChange?.(event as React.SyntheticEvent<HTMLElement>, {
        open: false,
      })
    }
    modalProps={{ isBlocking: false }}
  >
    {children}
  </FluentDialog>
);

const DialogSurface: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div>{children}</div>;
const DialogBody: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div>{children}</div>
);
const DialogTitle: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div className={styles.f2DialogTitle}>{children}</div>;
const DialogContent: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div>{children}</div>;
const DialogActions: React.FC<{ children?: React.ReactNode }> = ({
  children,
}) => <div className={styles.f2DialogActions}>{children}</div>;

const MultiLanguageContentEditor: React.FC<{
  content: ILanguageContent[];
  onContentChange: (content: ILanguageContent[]) => void;
  availableLanguages: ISupportedLanguage[];
  errors?: { [key: string]: string | undefined };
  linkUrl?: string;
  tenantDefaultLanguage?: TargetLanguage;
  context?: ApplicationCustomizerContext;
  imageFolderName?: string;
  disableImageUpload?: boolean;
  languagePolicy?: ILanguagePolicy;
  copilotService?: CopilotService;
}> = ({
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
  
  // Use ref to always access latest content (avoid stale closure in async callbacks)
  const contentRef = React.useRef(content);
  React.useEffect(() => {
    contentRef.current = content;
  }, [content]);
  
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
  }, [content, content.length, selectedTab]);

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

  const handleTranslate = React.useCallback(
    async (
      targetLanguage: string,
      targetLangName: string,
      overwriteExisting: boolean = true,
    ): Promise<void> => {
      const currentContent = contentRef.current;
      
      // 3. Fall back to any language with content (even if it's the target - for copy/paste scenarios)
      const sourceContent = 
        currentContent.find((c) => c.language === fallbackLanguage && (c.title || c.description)) ||
        currentContent.find((c) => c.language !== targetLanguage && (c.title || c.description)) ||
        currentContent.find((c) => c.title || c.description);
      
      const targetContentIndex = currentContent.findIndex(
        (c) => c.language === targetLanguage,
      );
      const targetContent =
        targetContentIndex >= 0 ? currentContent[targetContentIndex] : null;

      if (!sourceContent) {
        setTranslationError(
          "No source content found. Please add content to at least one language first.",
        );
        return;
      }

      if (
        targetContent &&
        !overwriteExisting &&
        (targetContent.title?.trim() || targetContent.description?.trim())
      ) {
        return;
      }

      setTranslatingLanguages((prev) => [...prev, targetLanguage]);
      setTranslationError(null);

      if (!copilotService) {
        setTranslationError("Copilot service not available");
        setTranslatingLanguages((prev) => prev.filter((l) => l !== targetLanguage));
        return;
      }

      try {
        const results: Partial<ILanguageContent> = {
          language: targetLanguage as TargetLanguage,
        };

        if (sourceContent.title) {
          const titleResponse = await copilotService.translateText(
            sourceContent.title,
            targetLangName,
          );
          if (titleResponse.isError) {
            throw new Error(
              titleResponse.errorMessage || "Title translation failed",
            );
          }
          results.title = titleResponse.content.trim();
        }

        if (sourceContent.description) {
          const descResponse = await copilotService.translateText(
            sourceContent.description,
            targetLangName,
          );
          if (descResponse.isError) {
            throw new Error(
              descResponse.errorMessage || "Description translation failed",
            );
          }
          results.description = descResponse.content.trim();
        }

        if (sourceContent.linkDescription) {
          const linkResponse = await copilotService.translateText(
            sourceContent.linkDescription,
            targetLangName,
          );
          if (linkResponse.isError) {
            throw new Error(
              linkResponse.errorMessage || "Link description translation failed",
            );
          }
          results.linkDescription = linkResponse.content.trim();
        }

        const updatedContent = currentContent.map((c) =>
          c.language === targetLanguage ? { ...c, ...results } : c,
        );
        onContentChange(updatedContent);

        logger.info(
          "MultiLanguageContentEditor",
          `Translated content to ${targetLangName}`,
        );
      } catch (error) {
        logger.error("MultiLanguageContentEditor", "Translation failed", error);
        setTranslationError(strings.CopilotTranslationFailed);
      } finally {
        setTranslatingLanguages((prev) => prev.filter((l) => l !== targetLanguage));
      }
    },
    [fallbackLanguage, copilotService, onContentChange, strings],
  );

  const handleTranslateAllMissing = React.useCallback(async (): Promise<void> => {
    const currentContent = contentRef.current;
    
    logger.debug("MultiLanguageContentEditor", "Translate all missing clicked", {
      contentCount: currentContent.length,
      languages: currentContent.map(c => ({ lang: c.language, hasTitle: !!c.title, hasDesc: !!c.description }))
    });
    
    const sourceContent = currentContent.find((c) => c.title || c.description);
    if (!sourceContent) {
      logger.debug("MultiLanguageContentEditor", "No source content found");
      setTranslationError(
        "No source content found. Please add content to at least one language first.",
      );
      return;
    }

    const missingTranslations = currentContent.filter(
      (c) =>
        c.language !== sourceContent.language &&
        (!c.title || !c.description),
    );

    logger.debug("MultiLanguageContentEditor", "Missing translations found", {
      sourceLanguage: sourceContent.language,
      missingCount: missingTranslations.length,
      missingLanguages: missingTranslations.map(c => c.language)
    });

    if (missingTranslations.length === 0) {
      setTranslationInfo("All languages already have content. Nothing to translate.");
      return;
    }

    setIsTranslatingAll(true);
    setTranslationError(null);

    try {
      for (const targetContent of missingTranslations) {
        const langInfo = getLanguageInfo(targetContent.language as TargetLanguage);
        if (langInfo) {
          await handleTranslate(targetContent.language, langInfo.nativeName, false);
        }
      }
    } catch (error) {
      logger.error("MultiLanguageContentEditor", "Batch translation failed", error);
      setTranslationError(strings.CopilotTranslationFailed);
    } finally {
      setIsTranslatingAll(false);
    }
  }, [handleTranslate, strings]);

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
                <Add24Regular className={styles.languageAddIcon} />
              </button>
            ))}
          </div>
          {copilotService && content.length > 1 && (
            <div className={styles.translationActions}>
              <Button
                appearance="secondary"
                size="small"
                icon={
                  isTranslatingAll ? (
                    <Spinner size="tiny" />
                  ) : (
                    <SparkleRegular />
                  )
                }
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
                    // Show translate button if there's content in another language to use as source
                    content.some(
                      (c) =>
                        c.language !== contentItem.language &&
                        (c.title?.trim() || c.description?.trim()),
                    ) && (
                      <div className={styles.translationRow}>
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
                            if (contentItem.title || contentItem.description) {
                              setConfirmOverwriteLang(contentItem.language);
                              return;
                            }

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
