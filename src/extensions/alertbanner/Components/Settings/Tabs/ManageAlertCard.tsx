import * as React from "react";
import { Delete24Regular, Edit24Regular, Globe24Regular } from "@fluentui/react-icons";
import { SharePointButton } from "../../UI/SharePointControls";
import {
  ContentStatus,
  ContentType,
  IAlertItem,
  TargetLanguage,
} from "../../Alerts/IAlerts";
import { ISupportedLanguage } from "../../Services/LanguageAwarenessService";
import { htmlSanitizer } from "../../Utils/HtmlSanitizer";
import { StringUtils } from "../../Utils/StringUtils";
import { DateUtils } from "../../Utils/DateUtils";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import styles from "../AlertSettings.module.scss";

export interface IDisplayAlertItem extends IAlertItem {
  isLanguageGroup?: boolean;
  languageVariants?: IAlertItem[];
}

export interface IManageAlertCardProps {
  alert: IDisplayAlertItem;
  isSelected: boolean;
  isBusy?: boolean;
  onSelectionChange: (checked: boolean) => void;
  onEdit: () => void;
  onDelete: () => void;
  onPublishDraft: () => void;
  onSubmitForReview: () => void;
  onApprove: () => void;
  onReject: () => void;
  priorityLabel: string;
  alertTypeStyle?: {
    backgroundColor?: string;
    textColor?: string;
  };
  supportedLanguageMap: Map<string, ISupportedLanguage>;
}

const ManageAlertCard: React.FC<IManageAlertCardProps> = ({
  alert,
  isSelected,
  isBusy = false,
  onSelectionChange,
  onEdit,
  onDelete,
  onPublishDraft,
  onSubmitForReview,
  onApprove,
  onReject,
  priorityLabel,
  alertTypeStyle,
  supportedLanguageMap,
}) => {
  const isMultiLanguage =
    !!alert.isLanguageGroup &&
    !!alert.languageVariants &&
    alert.languageVariants.length > 1;

  const languageSummary = React.useMemo(() => {
    if (!isMultiLanguage) {
      if (alert.targetLanguage === TargetLanguage.All) {
        return strings.ManageAlertsAllLanguagesLabel;
      }

      const lang = supportedLanguageMap.get(alert.targetLanguage);
      return `${lang?.flag || ""} ${lang?.nativeName || alert.targetLanguage}`;
    }

    const variants = alert.languageVariants || [];
    const flags = variants
      .map((variant) => supportedLanguageMap.get(variant.targetLanguage)?.flag || variant.targetLanguage)
      .join(" ");

    return Text.format(strings.ManageAlertsMultiLanguageList, flags);
  }, [
    alert.targetLanguage,
    alert.languageVariants,
    isMultiLanguage,
    supportedLanguageMap,
  ]);

  return (
    <div
      className={`${styles.alertCard} ${isSelected ? styles.selected : ""} ${alert.contentType === ContentType.Template ? styles.templateCard : ""}`}
    >
      <div className={styles.alertCardHeader}>
        <input
          type="checkbox"
          checked={isSelected}
          disabled={isBusy}
          aria-label={Text.format(
            strings.ManageAlertsSelectItemLabel,
            alert.title || alert.id,
          )}
          onChange={(e) => onSelectionChange(e.target.checked)}
          className={styles.alertCheckbox}
        />
        <div className={styles.alertStatus}>
          <span
            className={`${styles.statusBadge} ${alert.status.toLowerCase() === "active" ? styles.active : alert.status.toLowerCase() === "expired" ? styles.expired : styles.scheduled}`}
          >
            {alert.status}
          </span>
          {alert.isPinned && (
            <span className={styles.pinnedBadge}>{strings.ManageAlertsPinnedBadge}</span>
          )}
          {alert.contentStatus && alert.contentStatus !== ContentStatus.Approved && (
            <span className={`${styles.statusBadge} ${styles.template}`}>
              {alert.contentStatus === ContentStatus.PendingReview
                ? strings.ManageAlertsContentStatusPending
                : alert.contentStatus === ContentStatus.Rejected
                  ? strings.ManageAlertsContentStatusRejected
                  : strings.ManageAlertsContentStatusDraft}
            </span>
          )}
        </div>
      </div>

      <div
        className={styles.alertCardContent}
        style={
          alert.contentType === ContentType.Template
            ? {
                display: "block",
                visibility: "visible",
                opacity: 1,
              }
            : {}
        }
      >
        {alert.AlertType ? (
          <div
            className={styles.alertTypeIndicator}
            style={
              {
                "--bg-color": alertTypeStyle?.backgroundColor || "#0078d4",
                "--text-color": alertTypeStyle?.textColor || "#ffffff",
              } as React.CSSProperties
            }
          >
            {alert.AlertType}
          </div>
        ) : (
          <div
            className={styles.alertTypeIndicator}
            style={
              {
                "--bg-color":
                  alert.contentType === ContentType.Template ? "#8764b8" : "#0078d4",
                "--text-color": "#ffffff",
              } as React.CSSProperties
            }
          >
            {alert.contentType === ContentType.Template
              ? strings.ManageAlertsTemplateBadge
              : strings.ManageAlertsAlertBadge}
          </div>
        )}

        <h4 className={styles.alertCardTitle}>
          {alert.title || strings.ManageAlertsNoTitleFallback}
          {isMultiLanguage && (
            <span className={styles.multiLanguageBadge}>
              <Globe24Regular
                style={{
                  width: "12px",
                  height: "12px",
                  marginRight: "4px",
                }}
              />
              {Text.format(
                strings.ManageAlertsLanguageBadge,
                alert.languageVariants?.length || 0,
              )}
            </span>
          )}
        </h4>

        <div className={styles.alertCardDescription}>
          {alert.description ? (
            <div
              dangerouslySetInnerHTML={{
                __html: htmlSanitizer.sanitizeHtml(
                  StringUtils.truncate(alert.description, 150),
                ),
              }}
            />
          ) : (
            <em style={{ color: "#999" }}>{strings.ManageAlertsNoDescriptionFallback}</em>
          )}
        </div>

        <div className={styles.alertMetaData}>
          <div className={styles.metaInfo}>
            <strong>{strings.AlertType}:</strong>
            <span
              className={`${styles.contentTypeBadge} ${alert.contentType === ContentType.Template ? styles.template : styles.alert}`}
            >
              {alert.contentType === ContentType.Template
                ? strings.ManageAlertsTemplateBadge
                : strings.ManageAlertsAlertBadge}
            </span>
          </div>
          <div className={styles.metaInfo}>
            <strong>{strings.Priority}:</strong> {priorityLabel}
          </div>
          <div className={styles.metaInfo}>
            <strong>{strings.Language}:</strong>
            {isMultiLanguage ? (
              <span className={styles.languageList}>{languageSummary}</span>
            ) : (
              languageSummary
            )}
          </div>
          {alert.linkUrl && (
            <div className={styles.metaInfo}>
              <strong>{strings.LinkDescription}:</strong> {alert.linkDescription}
            </div>
          )}
          <div className={styles.metaInfo}>
            <strong>{strings.ManageAlertsCreatedLabel}:</strong>{" "}
            {DateUtils.formatForDisplay(alert.createdDate || new Date())}
          </div>
          {alert.scheduledStart && (
            <div className={styles.metaInfo}>
              <strong>{strings.CreateAlertStartDateLabel}:</strong>{" "}
              {DateUtils.formatDateTimeForDisplay(alert.scheduledStart)}
            </div>
          )}
          {alert.scheduledEnd && (
            <div className={styles.metaInfo}>
              <strong>{strings.CreateAlertEndDateLabel}:</strong>{" "}
              {DateUtils.formatDateTimeForDisplay(alert.scheduledEnd)}
            </div>
          )}
        </div>
      </div>

      <div className={styles.alertCardActions}>
        {alert.contentType === ContentType.Draft && (
          <SharePointButton
            variant="primary"
            title={strings.ManageAlertsPublishButtonLabel}
            onClick={onPublishDraft}
            disabled={isBusy}
          >
            {strings.ManageAlertsPublishButtonLabel}
          </SharePointButton>
        )}
        {alert.contentType !== ContentType.Draft &&
          (alert.contentStatus === ContentStatus.Draft ||
            alert.contentStatus === ContentStatus.Rejected ||
            !alert.contentStatus) && (
          <SharePointButton
            variant="primary"
            title={strings.ManageAlertsSubmitForReviewTitle}
            onClick={onSubmitForReview}
            disabled={isBusy}
          >
            {strings.ManageAlertsSubmitForReviewButton}
          </SharePointButton>
          )}
        {alert.contentStatus === ContentStatus.PendingReview && (
          <>
            <SharePointButton variant="primary" onClick={onApprove} disabled={isBusy}>
              {strings.ManageAlertsApproveButton}
            </SharePointButton>
            <SharePointButton variant="danger" onClick={onReject} disabled={isBusy}>
              {strings.ManageAlertsRejectButton}
            </SharePointButton>
          </>
        )}

        <SharePointButton
          variant="secondary"
          icon={<Edit24Regular />}
          onClick={onEdit}
          disabled={isBusy}
        >
          {strings.Edit}
        </SharePointButton>
        <SharePointButton
          variant="danger"
          icon={<Delete24Regular />}
          onClick={onDelete}
          disabled={isBusy}
        >
          {strings.Delete}
        </SharePointButton>
      </div>
    </div>
  );
};

export default React.memo(ManageAlertCard);
