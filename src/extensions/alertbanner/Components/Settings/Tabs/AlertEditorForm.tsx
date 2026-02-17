import * as React from "react";
import { Eye24Regular } from "@fluentui/react-icons";
import {
  ISharePointSelectOption,
  SharePointButton,
  SharePointInput,
  SharePointPeoplePicker,
  SharePointSection,
  SharePointSelect,
  SharePointToggle,
} from "../../UI/SharePointControls";
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CopilotDraftControl } from "../../CopilotControls/CopilotDraftControl";
import { CopilotGovernanceControl } from "../../CopilotControls/CopilotGovernanceControl";
import SharePointRichTextEditor from "../../UI/SharePointRichTextEditor";
import MultiLanguageContentEditor from "../../UI/MultiLanguageContentEditor";
import AlertPreview from "../../UI/AlertPreview";
import SiteSelector from "../../UI/SiteSelector";
import {
  AlertPriority,
  ContentType,
  IAlertType,
  IPersonField,
  NotificationType,
  TargetLanguage,
} from "../../Alerts/IAlerts";
import {
  ILanguageContent,
  ISupportedLanguage,
} from "../../Services/LanguageAwarenessService";
import { ILanguagePolicy } from "../../Services/LanguagePolicyService";
import { SiteContextDetector } from "../../Utils/SiteContextDetector";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { NotificationService } from "../../Services/NotificationService";
import { CopilotService } from "../../Services/CopilotService";
import { DateUtils } from "../../Utils/DateUtils";
import { Text } from "@microsoft/sp-core-library";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import styles from "../AlertSettings.module.scss";
import { IFormErrors } from "./SharedTypes";
import { validateAlertData } from "../../Utils/AlertValidation";
import { getLocalizedValidationMessage } from "../../Utils/AlertValidationLocalization";

export interface IAlertEditorState {
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  notificationType: NotificationType;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string[];
  scheduledStart?: Date;
  scheduledEnd?: Date;
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageContent?: ILanguageContent[];
  languageGroup?: string;
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];
}

export type CreateWizardStep = "content" | "audience" | "publish";

const stripRichText = (value: string): string =>
  value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();

export interface IAlertEditorFormProps<T extends IAlertEditorState> {
  mode: "create" | "manage";
  alert: T;
  setAlert: React.Dispatch<React.SetStateAction<T>>;
  errors: IFormErrors;
  setErrors: React.Dispatch<React.SetStateAction<IFormErrors>>;
  alertTypes: IAlertType[];
  alertTypeOptions: ISharePointSelectOption[];
  priorityOptions: ISharePointSelectOption[];
  notificationOptions: ISharePointSelectOption[];
  contentTypeOptions: ISharePointSelectOption[];
  languageOptions: ISharePointSelectOption[];
  supportedLanguages: ISupportedLanguage[];
  useMultiLanguage: boolean;
  setUseMultiLanguage: React.Dispatch<React.SetStateAction<boolean>>;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  siteDetector: SiteContextDetector;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  languagePolicy: ILanguagePolicy;
  tenantDefaultLanguage: TargetLanguage;
  copilotEnabled?: boolean;
  copilotService: CopilotService;
  notificationService: NotificationService;
  isBusy: boolean;
  showPreview: boolean;
  setShowPreview: React.Dispatch<React.SetStateAction<boolean>>;
  imageFolderName?: string;
  disableImageUpload?: boolean;
  renderMiddleSections?: React.ReactNode;
  actionButtons: React.ReactNode;
  afterActions?: React.ReactNode;
  showScheduleSummary?: boolean;
  applySelectedAlertTypeDefaultPriority?: boolean;
  createStep?: CreateWizardStep;
  onCreateStepChange?: (step: CreateWizardStep) => void;
  stepCompletion?: Record<CreateWizardStep, boolean>;
  copilotAvailability?: "unknown" | "available" | "unavailable";
  copilotAvailabilityMessage?: string;
}

const AlertEditorForm = <T extends IAlertEditorState>({
  mode,
  alert,
  setAlert,
  errors,
  setErrors,
  alertTypes,
  alertTypeOptions,
  priorityOptions,
  notificationOptions,
  contentTypeOptions,
  languageOptions,
  supportedLanguages,
  useMultiLanguage,
  setUseMultiLanguage,
  userTargetingEnabled,
  notificationsEnabled,
  enableTargetSite,
  siteDetector,
  graphClient,
  context,
  languagePolicy,
  tenantDefaultLanguage,
  copilotEnabled,
  copilotService,
  notificationService,
  isBusy,
  showPreview,
  setShowPreview,
  imageFolderName,
  disableImageUpload,
  renderMiddleSections,
  actionButtons,
  afterActions,
  showScheduleSummary,
  applySelectedAlertTypeDefaultPriority,
  createStep,
  onCreateStepChange,
  stepCompletion,
  copilotAvailability,
  copilotAvailabilityMessage,
}: IAlertEditorFormProps<T>): JSX.Element => {
  const languageContent = alert.languageContent || [];
  const editorRootRef = React.useRef<HTMLDivElement>(null);
  const wasBusyRef = React.useRef(isBusy);
  const previousAlertTypeRef = React.useRef<string>("");
  const previewRegionId = React.useMemo(
    () => `alert-preview-${Math.random().toString(36).slice(2, 10)}`,
    [],
  );
  const languageModeGroupId = React.useMemo(
    () => `language-mode-${Math.random().toString(36).slice(2, 10)}`,
    [],
  );

  React.useEffect(() => {
    if (wasBusyRef.current && !isBusy) {
      editorRootRef.current?.focus();
    }
    wasBusyRef.current = isBusy;
  }, [isBusy]);

  React.useEffect(() => {
    if (mode !== "create" || !applySelectedAlertTypeDefaultPriority) {
      previousAlertTypeRef.current = alert.AlertType || "";
      return;
    }

    const currentAlertType = alert.AlertType || "";
    const hasAlertTypeChanged = previousAlertTypeRef.current !== currentAlertType;
    previousAlertTypeRef.current = currentAlertType;

    if (!hasAlertTypeChanged || !currentAlertType) {
      return;
    }

    const selectedType = alertTypes.find((type) => type.name === currentAlertType);
    if (
      selectedType?.defaultPriority &&
      alert.priority !== selectedType.defaultPriority
    ) {
      setAlert((prev) =>
        ({ ...prev, priority: selectedType.defaultPriority } as T),
      );
    }
  }, [
    alert.AlertType,
    alert.priority,
    alertTypes,
    applySelectedAlertTypeDefaultPriority,
    mode,
    setAlert,
  ]);

  const handlePeoplePickerChange = React.useCallback(
    (items: any[]) => {
      const users: IPersonField[] = [];
      const groups: IPersonField[] = [];

      items.forEach((item) => {
        const personField: IPersonField = {
          id: item.id,
          displayName: item.text,
          email: item.secondaryText,
          loginName: item.loginName,
          isGroup: item.id ? item.id.indexOf("c:0(.s|true") === -1 : false,
        };

        if (
          item.imageInitials ||
          (item.secondaryText && item.secondaryText.indexOf("@") > -1)
        ) {
          personField.isGroup = false;
          users.push(personField);
        } else {
          personField.isGroup = true;
          groups.push(personField);
        }
      });

      setAlert((prev) =>
        ({
          ...prev,
          targetUsers: users,
          targetGroups: groups,
        }) as T,
      );
    },
    [setAlert],
  );

  const getCurrentAlertType = React.useCallback((): IAlertType | undefined => {
    return alertTypes.find((type) => type.name === alert.AlertType);
  }, [alertTypes, alert.AlertType]);

  const languageConfigurationLabel =
    mode === "manage"
      ? strings.ManageAlertsLanguageConfigurationLabel
      : strings.CreateAlertLanguageConfigurationLabel;

  const singleLanguageButtonLabel =
    mode === "manage"
      ? strings.ManageAlertsSingleLanguageButton
      : strings.CreateAlertSingleLanguageButton;

  const multiLanguageButtonLabel =
    mode === "manage"
      ? strings.ManageAlertsMultiLanguageButton
      : strings.CreateAlertMultiLanguageButton;

  const contentClassificationDescription =
    mode === "manage"
      ? strings.ManageAlertsContentClassificationDescription
      : strings.CreateAlertSectionContentClassificationDescription;

  const targetLanguageDescription =
    mode === "manage"
      ? strings.ManageAlertsTargetLanguageDescription
      : strings.CreateAlertSectionLanguageTargetingDescription;

  const descriptionPlaceholder =
    mode === "manage"
      ? strings.ManageAlertsDescriptionPlaceholder
      : strings.CreateAlertDescriptionPlaceholder;

  const descriptionHelp =
    mode === "manage"
      ? strings.ManageAlertsDescriptionHelp
      : strings.CreateAlertDescriptionHelp;

  const alertTypeDescription =
    mode === "manage"
      ? strings.ManageAlertsAlertTypeDescription
      : strings.CreateAlertConfigurationDescription;

  const priorityDescription =
    mode === "manage"
      ? strings.ManageAlertsPriorityDescription
      : strings.CreateAlertPriorityDescription;

  const pinDescription =
    mode === "manage"
      ? strings.ManageAlertsPinDescription
      : strings.CreateAlertPinDescription;

  const notificationDescription =
    mode === "manage"
      ? strings.ManageAlertsNotificationDescription
      : strings.CreateAlertNotificationDescription;

  const notificationLabel =
    mode === "manage"
      ? strings.ManageAlertsNotificationMethodLabel
      : strings.CreateAlertNotificationLabel;

  const linkDescriptionInfo =
    mode === "manage"
      ? strings.ManageAlertsLinkDescriptionInfo
      : strings.CreateAlertLinkDescriptionInfo;

  const startDateDescription =
    mode === "manage"
      ? strings.ManageAlertsStartDateDescription
      : strings.CreateAlertStartDateDescription;

  const endDateDescription =
    mode === "manage"
      ? strings.ManageAlertsEndDateDescription
      : strings.CreateAlertEndDateDescription;

  const isCreateMode = mode === "create";
  const currentCreateStep: CreateWizardStep = isCreateMode
    ? createStep || "content"
    : "content";
  const canUseCopilot = !!copilotEnabled && copilotAvailability === "available";
  const shouldShowCopilotNotice =
    !!copilotEnabled &&
    copilotAvailability === "unavailable" &&
    !!copilotAvailabilityMessage;

  const wizardSteps = React.useMemo(
    () => [
      {
        id: "content" as CreateWizardStep,
        label: strings.CreateAlertWizardContentStep,
      },
      {
        id: "audience" as CreateWizardStep,
        label: strings.CreateAlertWizardAudienceStep,
      },
      {
        id: "publish" as CreateWizardStep,
        label: strings.CreateAlertWizardPublishStep,
      },
    ],
    [],
  );

  const shouldRenderStep = React.useCallback(
    (step: CreateWizardStep): boolean =>
      !isCreateMode || currentCreateStep === step,
    [currentCreateStep, isCreateMode],
  );

  const getValidationErrors = React.useCallback(
    (nextAlert: T): IFormErrors =>
      validateAlertData(nextAlert, {
        useMultiLanguage,
        languagePolicy,
        tenantDefaultLanguage,
        getString: getLocalizedValidationMessage,
        validateTargetSites: enableTargetSite,
      }),
    [
      enableTargetSite,
      languagePolicy,
      tenantDefaultLanguage,
      useMultiLanguage,
    ],
  );

  const applyFieldValidation = React.useCallback(
    (nextAlert: T, fields: string[]) => {
      const nextErrors = getValidationErrors(nextAlert);

      if (!useMultiLanguage && fields.includes("description")) {
        const plainDescription = stripRichText(nextAlert.description || "");
        if (plainDescription.length === 0) {
          nextErrors.description = getLocalizedValidationMessage(
            "DescriptionRequired",
          );
        } else if (plainDescription.length < 10) {
          nextErrors.description = getLocalizedValidationMessage(
            "DescriptionMinLength",
          );
        } else {
          delete nextErrors.description;
        }
      }

      if (!useMultiLanguage && fields.includes("title")) {
        const plainTitle = (nextAlert.title || "").trim();
        if (plainTitle.length === 0) {
          nextErrors.title = getLocalizedValidationMessage("TitleRequired");
        } else if (plainTitle.length < 3) {
          nextErrors.title = getLocalizedValidationMessage("TitleMinLength");
        } else {
          delete nextErrors.title;
        }
      }

      setErrors((prev) => {
        const updated: IFormErrors = { ...prev };
        fields.forEach((field) => {
          if (nextErrors[field]) {
            updated[field] = nextErrors[field];
          } else {
            delete updated[field];
          }
        });
        return updated;
      });
    },
    [getValidationErrors, setErrors],
  );

  const updateAlertFields = React.useCallback(
    (patch: Partial<T>, validateFields?: string[]) => {
      setAlert((prev) => {
        const next = { ...prev, ...patch } as T;
        if (validateFields && validateFields.length > 0) {
          applyFieldValidation(next, validateFields);
        }
        return next;
      });
    },
    [applyFieldValidation, setAlert],
  );

  const showPriorityInheritanceMessage =
    mode === "create" &&
    !!applySelectedAlertTypeDefaultPriority &&
    !!alertTypes.find(
      (alertType) =>
        alertType.name === alert.AlertType && !!alertType.defaultPriority,
    );

  return (
    <div
      className={`${styles.alertForm} ${isCreateMode ? styles.createEditorForm : ""}`}
      ref={editorRootRef}
      tabIndex={-1}
      aria-label={
        mode === "manage"
          ? strings.ManageAlerts
          : strings.CreateAlert
      }
    >
      <div className={styles.formWithPreview}>
        <div className={styles.formColumn}>
          {isCreateMode && (
            <div className={styles.createWizardSteps}>
              {wizardSteps.map((step) => {
                const isActive = currentCreateStep === step.id;
                const isComplete = !!stepCompletion?.[step.id];
                return (
                  <button
                    key={step.id}
                    type="button"
                    className={`${styles.createWizardStep} ${isActive ? styles.activeWizardStep : ""}`}
                    onClick={() => onCreateStepChange?.(step.id)}
                    aria-current={isActive ? "step" : undefined}
                  >
                    <span className={styles.createWizardStepLabel}>
                      {step.label}
                    </span>
                    <span
                      className={`${styles.createWizardStepStatus} ${isComplete ? styles.wizardStepComplete : styles.wizardStepIncomplete}`}
                    >
                      {isComplete
                        ? strings.CreateAlertWizardStepComplete
                        : strings.CreateAlertWizardStepIncomplete}
                    </span>
                  </button>
                );
              })}
            </div>
          )}

          {shouldRenderStep("content") && (
            <SharePointSection
              title={strings.CreateAlertSectionContentClassificationTitle}
            >
              <SharePointSelect
                label={strings.ContentTypeLabel}
                value={alert.contentType}
                onChange={(value) =>
                  updateAlertFields({
                    contentType: value as ContentType,
                  } as Partial<T>)
                }
                options={contentTypeOptions}
                required
                description={contentClassificationDescription}
              />

              <div className={styles.languageModeSelector}>
                <label className={styles.fieldLabel}>
                  {languageConfigurationLabel}
                </label>
                <div
                  id={languageModeGroupId}
                  className={styles.languageOptions}
                  role="group"
                  aria-label={languageConfigurationLabel}
                >
                  <SharePointButton
                    variant={!useMultiLanguage ? "primary" : "secondary"}
                    onClick={() => setUseMultiLanguage(false)}
                    aria-pressed={!useMultiLanguage}
                  >
                    {singleLanguageButtonLabel}
                  </SharePointButton>
                  <SharePointButton
                    variant={useMultiLanguage ? "primary" : "secondary"}
                    onClick={() => setUseMultiLanguage(true)}
                    aria-pressed={useMultiLanguage}
                  >
                    {multiLanguageButtonLabel}
                  </SharePointButton>
                </div>
              </div>
            </SharePointSection>
          )}

          {userTargetingEnabled && shouldRenderStep("audience") && (
            <SharePointSection title={strings.CreateAlertSectionUserTargetingTitle}>
              <SharePointPeoplePicker
                context={context}
                titleText={strings.CreateAlertPeoplePickerLabel}
                personSelectionLimit={50}
                groupName={""}
                showtooltip={true}
                required={false}
                disabled={isBusy}
                onChange={handlePeoplePickerChange}
                defaultSelectedUsers={[
                  ...(alert.targetUsers?.map((u) => u.email || u.loginName || u.displayName) || []),
                  ...(alert.targetGroups?.map((g) => g.loginName || g.displayName) || []),
                ]}
                principalTypes={[
                  PrincipalType.User,
                  PrincipalType.SharePointGroup,
                  PrincipalType.SecurityGroup,
                  PrincipalType.DistributionList,
                ]}
                description={strings.CreateAlertPeoplePickerDescription}
              />
            </SharePointSection>
          )}

          {shouldRenderStep("content") &&
            (!useMultiLanguage ? (
            <>
              <SharePointSection title={strings.CreateAlertSectionLanguageTargetingTitle}>
                <SharePointSelect
                  label={strings.CreateAlertTargetLanguageLabel}
                  value={alert.targetLanguage}
                  onChange={(value) =>
                    updateAlertFields({
                      targetLanguage: value as TargetLanguage,
                    } as Partial<T>)
                  }
                  options={languageOptions}
                  required
                  description={targetLanguageDescription}
                />
              </SharePointSection>

              <SharePointSection title={strings.CreateAlertSectionBasicInformationTitle}>
                <SharePointInput
                  label={strings.AlertTitle}
                  value={alert.title}
                  onChange={(value) => {
                    updateAlertFields(
                      { title: value } as Partial<T>,
                      ["title"],
                    );
                  }}
                  placeholder={strings.CreateAlertTitlePlaceholder}
                  required
                  error={errors.title}
                  description={strings.CreateAlertTitleDescription}
                />

                {canUseCopilot && (
                  <div className={styles.copilotActionsRow}>
                    <CopilotDraftControl
                      copilotService={copilotService}
                      onDraftGenerated={(draft) =>
                        updateAlertFields(
                          { description: draft } as Partial<T>,
                          ["description"],
                        )
                      }
                      onError={(error) =>
                        notificationService.showError(error, strings.CopilotErrorTitle)
                      }
                      disabled={isBusy}
                    />
                  </div>
                )}

                <SharePointRichTextEditor
                  label={strings.AlertDescription}
                  value={alert.description}
                  onChange={(value) => {
                    updateAlertFields(
                      { description: value } as Partial<T>,
                      ["description"],
                    );
                  }}
                  context={context}
                  placeholder={descriptionPlaceholder}
                  required
                  error={errors.description}
                  description={descriptionHelp}
                  imageFolderName={imageFolderName || alert.languageGroup || alert.title || "Untitled_Alert"}
                  disableImageUpload={disableImageUpload}
                />

                {canUseCopilot && (
                  <CopilotGovernanceControl
                    copilotService={copilotService}
                    textToAnalyze={alert.description}
                    onError={(error) =>
                      notificationService.showError(error, strings.CopilotErrorTitle)
                    }
                    disabled={isBusy}
                  />
                )}
              </SharePointSection>
            </>
          ) : (
            <SharePointSection title={strings.MultiLanguageContent}>
              <MultiLanguageContentEditor
                content={languageContent}
                onContentChange={(content) =>
                  setAlert((prev) => ({ ...prev, languageContent: content }) as T)
                }
                availableLanguages={supportedLanguages}
                errors={errors}
                linkUrl={alert.linkUrl || ""}
                context={context}
                imageFolderName={alert.languageGroup}
                disableImageUpload={disableImageUpload}
                tenantDefaultLanguage={tenantDefaultLanguage}
                languagePolicy={languagePolicy}
                copilotService={canUseCopilot ? copilotService : undefined}
              />
            </SharePointSection>
          ))}

          {shouldShowCopilotNotice && shouldRenderStep("content") && (
            <div className={styles.infoMessage}>
              <p>{copilotAvailabilityMessage}</p>
            </div>
          )}

          {shouldRenderStep("audience") && renderMiddleSections}

          {shouldRenderStep("content") && (
            <SharePointSection title={strings.CreateAlertConfigurationSectionTitle}>
            <SharePointSelect
              label={strings.AlertType}
              value={alert.AlertType}
              onChange={(value) => {
                const selectedType = alertTypes.find((t) => t.name === value);
                updateAlertFields(
                  {
                    AlertType: value,
                    priority:
                      applySelectedAlertTypeDefaultPriority &&
                      selectedType?.defaultPriority
                        ? selectedType.defaultPriority
                        : alert.priority,
                  } as Partial<T>,
                  ["AlertType"],
                );
              }}
              options={alertTypeOptions}
              required
              error={errors.AlertType}
              description={alertTypeDescription}
            />

            <SharePointSelect
              label={strings.CreateAlertPriorityLabel}
              value={alert.priority}
              onChange={(value) =>
                updateAlertFields({
                  priority: value as AlertPriority,
                } as Partial<T>)
              }
              options={priorityOptions}
              required
              description={priorityDescription}
            />

            {showPriorityInheritanceMessage && (
              <div className={styles.infoMessage}>
                <p>
                  {Text.format(
                    strings.CreateAlertPriorityInheritedMessage,
                    alert.AlertType,
                  )}
                </p>
              </div>
            )}

            <SharePointToggle
              label={strings.CreateAlertPinLabel}
              checked={alert.isPinned}
              onChange={(checked) =>
                updateAlertFields({ isPinned: checked } as Partial<T>)
              }
              description={pinDescription}
            />

            {notificationsEnabled && (
              <SharePointSelect
                label={notificationLabel}
                value={alert.notificationType}
                onChange={(value) =>
                  updateAlertFields({
                    notificationType: value as NotificationType,
                  } as Partial<T>)
                }
                options={notificationOptions}
                description={notificationDescription}
              />
            )}
            </SharePointSection>
          )}

          {shouldRenderStep("audience") && (
            <SharePointSection title={strings.CreateAlertActionLinkSectionTitle}>
            <SharePointInput
              label={strings.CreateAlertLinkUrlLabel}
              value={alert.linkUrl || ""}
              onChange={(value) => {
                updateAlertFields(
                  {
                    linkUrl: value,
                    linkDescription: value ? alert.linkDescription : "",
                  } as Partial<T>,
                  ["linkUrl", "linkDescription"],
                );
              }}
              placeholder={strings.CreateAlertLinkUrlPlaceholder}
              error={errors.linkUrl}
              description={strings.CreateAlertLinkUrlDescription}
            />

            {alert.linkUrl && !useMultiLanguage && (
                <SharePointInput
                  label={strings.CreateAlertLinkDescriptionLabel}
                  value={alert.linkDescription || ""}
                  onChange={(value) => {
                    updateAlertFields(
                      { linkDescription: value } as Partial<T>,
                      ["linkDescription"],
                    );
                  }}
                  placeholder={strings.CreateAlertLinkDescriptionPlaceholder}
                required={!!alert.linkUrl}
                error={errors.linkDescription}
                description={strings.CreateAlertLinkDescriptionDescription}
              />
            )}

            {alert.linkUrl && useMultiLanguage && (
              <div className={styles.infoMessage}>
                <p>{linkDescriptionInfo}</p>
              </div>
            )}
            </SharePointSection>
          )}

          {enableTargetSite && shouldRenderStep("audience") && (
            <SharePointSection title={strings.CreateAlertTargetSitesSectionTitle}>
              <SiteSelector
                selectedSites={alert.targetSites || []}
                onSitesChange={(sites) => {
                  updateAlertFields(
                    { targetSites: sites } as Partial<T>,
                    ["targetSites"],
                  );
                }}
                siteDetector={siteDetector}
                graphClient={graphClient}
                showPermissionStatus={mode === "create"}
                className={styles.siteTargeting}
              />
              {errors.targetSites && (
                <div className={styles.errorMessage}>{errors.targetSites}</div>
              )}
              {mode === "manage" && (
                <div className={styles.fieldDescription}>
                  {strings.ManageAlertsTargetSitesDescription}
                </div>
              )}
            </SharePointSection>
          )}

          {shouldRenderStep("publish") && (
            <SharePointSection title={strings.CreateAlertSchedulingSectionTitle}>
            {mode === "manage" && (
              <div className={styles.schedulingHeader}>
                <p className={styles.schedulingDescription}>
                  {strings.ManageAlertsSchedulingDescription}
                </p>
              </div>
            )}

            {mode === "create" && (
              <div className={styles.schedulePresets}>
                <SharePointButton
                  variant="secondary"
                  onClick={() =>
                    updateAlertFields(
                      { scheduledStart: new Date() } as Partial<T>,
                      ["scheduledStart", "scheduledEnd"],
                    )
                  }
                >
                  {strings.CreateAlertSchedulePresetNow}
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={() => {
                    const endOfDay = new Date();
                    endOfDay.setHours(23, 59, 0, 0);
                    updateAlertFields(
                      { scheduledEnd: endOfDay } as Partial<T>,
                      ["scheduledStart", "scheduledEnd"],
                    );
                  }}
                >
                  {strings.CreateAlertSchedulePresetEndOfDay}
                </SharePointButton>
                <SharePointButton
                  variant="secondary"
                  onClick={() => {
                    const oneWeek = new Date();
                    oneWeek.setDate(oneWeek.getDate() + 7);
                    updateAlertFields(
                      { scheduledEnd: oneWeek } as Partial<T>,
                      ["scheduledStart", "scheduledEnd"],
                    );
                  }}
                >
                  {strings.CreateAlertSchedulePresetPlusWeek}
                </SharePointButton>
              </div>
            )}

            <SharePointInput
              label={strings.CreateAlertStartDateLabel}
              type="datetime-local"
              value={DateUtils.toDateTimeLocalValue(alert.scheduledStart)}
              onChange={(value) => {
                updateAlertFields(
                  {
                    scheduledStart: value ? new Date(value) : undefined,
                  } as Partial<T>,
                  ["scheduledStart", "scheduledEnd"],
                );
              }}
              error={errors.scheduledStart}
              description={startDateDescription}
            />

            <SharePointInput
              label={strings.CreateAlertEndDateLabel}
              type="datetime-local"
              value={DateUtils.toDateTimeLocalValue(alert.scheduledEnd)}
              onChange={(value) => {
                updateAlertFields(
                  {
                    scheduledEnd: value ? new Date(value) : undefined,
                  } as Partial<T>,
                  ["scheduledStart", "scheduledEnd"],
                );
              }}
              error={errors.scheduledEnd}
              description={endDateDescription}
            />

            {mode === "create" && (
              <div className={styles.timezoneInfo}>
                <p>
                  {Text.format(
                    strings.ManageAlertsScheduleTimezone,
                    Intl.DateTimeFormat().resolvedOptions().timeZone,
                  )}
                </p>
              </div>
            )}

            {showScheduleSummary && (
              <>
                <div className={styles.scheduleSummary}>
                  <h4>{strings.ManageAlertsScheduleSummaryTitle}</h4>
                  {!alert.scheduledStart && !alert.scheduledEnd ? (
                    <p>{strings.ManageAlertsScheduleImmediate}</p>
                  ) : alert.scheduledStart && !alert.scheduledEnd ? (
                    <p>
                      {Text.format(
                        strings.ManageAlertsScheduleStartOnly,
                        new Date(alert.scheduledStart).toLocaleString(),
                      )}
                    </p>
                  ) : !alert.scheduledStart && alert.scheduledEnd ? (
                    <p>
                      {Text.format(
                        strings.ManageAlertsScheduleEndOnly,
                        new Date(alert.scheduledEnd).toLocaleString(),
                      )}
                    </p>
                  ) : (
                    <p>
                      {Text.format(
                        strings.ManageAlertsScheduleWindow,
                        new Date(alert.scheduledStart!).toLocaleString(),
                        new Date(alert.scheduledEnd!).toLocaleString(),
                      )}
                    </p>
                  )}
                </div>
                <div className={styles.timezoneInfo}>
                  <p>
                    {Text.format(
                      strings.ManageAlertsScheduleTimezone,
                      Intl.DateTimeFormat().resolvedOptions().timeZone,
                    )}
                  </p>
                </div>
              </>
            )}
            </SharePointSection>
          )}

          <div
            className={`${styles.formActions} ${isCreateMode ? styles.createActionBar : ""}`}
          >
            <div className={styles.formActionsPrimary}>{actionButtons}</div>
            <div className={styles.formActionsUtility}>
              <SharePointButton
                variant="secondary"
                onClick={() => setShowPreview(!showPreview)}
                icon={<Eye24Regular />}
                aria-expanded={showPreview}
                aria-controls={previewRegionId}
              >
                {showPreview ? strings.CreateAlertHidePreview : strings.CreateAlertShowPreview}
              </SharePointButton>
            </div>
          </div>

          {afterActions && (
            <div role="status" aria-live="polite">
              {afterActions}
            </div>
          )}
        </div>

        {showPreview && (
          <div className={styles.previewColumn} id={previewRegionId}>
            <div className={styles.previewSticky}>
              <div className={styles.alertCard}>
                {mode === "manage" && <h3>{strings.ManageAlertsLivePreviewTitle}</h3>}

                {useMultiLanguage && languageContent.length > 0 && (
                  <div className={styles.previewLanguageSelector}>
                    <label className={styles.previewLabel}>
                      {strings.CreateAlertPreviewLanguageLabel}
                    </label>
                    <div className={styles.previewLanguageButtons}>
                      {languageContent.map((content, index) => {
                        const lang = supportedLanguages.find(
                          (l) => l.code === content.language,
                        );
                        return (
                          <SharePointButton
                            key={content.language}
                            variant={index === 0 ? "primary" : "secondary"}
                            aria-pressed={index === 0}
                            onClick={() => {
                              const reorderedContent = [
                                content,
                                ...languageContent.filter((_, i) => i !== index),
                              ];
                              setAlert((prev) =>
                                ({
                                  ...prev,
                                  languageContent: reorderedContent,
                                }) as T,
                              );
                            }}
                            className={styles.previewLanguageButton}
                          >
                            {lang?.flag || content.language} {lang?.nativeName || content.language}
                          </SharePointButton>
                        );
                      })}
                    </div>
                  </div>
                )}

                <AlertPreview
                  title={
                    useMultiLanguage && languageContent.length > 0
                      ? languageContent[0]?.title || strings.CreateAlertMultiLanguagePreviewTitle
                      : alert.title || strings.AlertPreviewDefaultTitle
                  }
                  description={
                    useMultiLanguage && languageContent.length > 0
                      ? languageContent[0]?.description || strings.CreateAlertMultiLanguagePreviewDescription
                      : alert.description || strings.AlertPreviewDefaultDescription
                  }
                  alertType={
                    getCurrentAlertType() || {
                      name: strings.AlertTypeInfo,
                      iconName: "Info",
                      backgroundColor: "#0078d4",
                      textColor: "#ffffff",
                      additionalStyles: "",
                      priorityStyles: {},
                    }
                  }
                  priority={alert.priority}
                  linkUrl={alert.linkUrl || ""}
                  linkDescription={
                    useMultiLanguage && languageContent.length > 0
                      ? languageContent[0]?.linkDescription || strings.CreateAlertLinkPreviewFallback
                      : alert.linkDescription || strings.CreateAlertLinkPreviewFallback
                  }
                  isPinned={alert.isPinned}
                />

                {useMultiLanguage && languageContent.length > 0 && (
                  <div className={styles.multiLanguagePreviewInfo} role="status" aria-live="polite">
                    <p>
                      <strong>
                        {mode === "manage"
                          ? strings.ManageAlertsMultiLanguagePreviewHeading
                          : strings.CreateAlertMultiLanguagePreviewHeading}
                      </strong>
                    </p>
                    <p>
                      {Text.format(
                        mode === "manage"
                          ? strings.ManageAlertsMultiLanguagePreviewCurrentLanguage
                          : strings.CreateAlertMultiLanguagePreviewCurrentLanguage,
                        supportedLanguages.find(
                          (l) => l.code === languageContent[0]?.language,
                        )?.nativeName || languageContent[0]?.language,
                      )}
                    </p>
                    <p>
                      {Text.format(
                        mode === "manage"
                          ? strings.ManageAlertsMultiLanguagePreviewAvailableLanguages
                          : strings.CreateAlertMultiLanguagePreviewAvailableLanguages,
                        languageContent.length,
                        languageContent
                          .map(
                            (c) =>
                              supportedLanguages.find((l) => l.code === c.language)?.flag || c.language,
                          )
                          .join(" "),
                      )}
                    </p>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default AlertEditorForm;
