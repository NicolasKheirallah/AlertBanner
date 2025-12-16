import * as React from "react";
import { StringUtils } from "../Utils/StringUtils";
import { IAlertItem } from "../Alerts/IAlerts";
import { Document24Regular, ArrowDownload24Regular } from "@fluentui/react-icons";
import DescriptionContent from "./DescriptionContent";
import styles from "./AlertItem.module.scss";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

interface IAlertContentProps {
  item: IAlertItem;
  expanded: boolean;
  stopPropagation: (e: React.MouseEvent) => void;
}

const AlertContent: React.FC<IAlertContentProps> = React.memo(({ item, expanded, stopPropagation }) => {
  if (!expanded) return null;

  const resolveAttachmentUrl = (serverRelativeUrl?: string): string => {
    if (!serverRelativeUrl) {
      return '#';
    }

    if (/^https?:\/\//i.test(serverRelativeUrl)) {
      return serverRelativeUrl;
    }

    if (typeof window === 'undefined') {
      return serverRelativeUrl;
    }

    return `${window.location.origin}${serverRelativeUrl}`;
  };

  const formatFileSize = (bytes?: number): string => {
    if (!bytes) return '';

    const kb = bytes / 1024;
    if (kb < 1024) {
      return Text.format(strings.FileSizeKilobytes, kb.toFixed(1));
    }

    const mb = kb / 1024;
    return Text.format(strings.FileSizeMegabytes, mb.toFixed(1));
  };

  return (
    <div
      className={styles.alertContentContainer}
      onClick={stopPropagation}
    >
      {item.description && <DescriptionContent description={item.description} isAlertExpanded={expanded} />}

      {item.attachments && item.attachments.length > 0 && (
        <div className={styles.attachmentsSection}>
          <div className={styles.attachmentsTitle}>
            {Text.format(strings.AttachmentsHeader, item.attachments.length.toString())}
          </div>
          <div className={styles.attachmentsList}>
            {item.attachments.map((attachment, index) => (
              <a
                key={index}
                href={StringUtils.resolveUrl(attachment.serverRelativeUrl)}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.attachmentItem}
                onClick={stopPropagation}
              >
                <Document24Regular className={styles.attachmentIcon} />
                <div className={styles.attachmentInfo}>
                  <div className={styles.attachmentName}>{attachment.fileName}</div>
                  {attachment.size && (
                    <div className={styles.attachmentSize}>{formatFileSize(attachment.size)}</div>
                  )}
                </div>
                <ArrowDownload24Regular className={styles.downloadIcon} />
              </a>
            ))}
          </div>
        </div>
      )}
    </div>
  );
});

AlertContent.displayName = 'AlertContent';

export default AlertContent;
