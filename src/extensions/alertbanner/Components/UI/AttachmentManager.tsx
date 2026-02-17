import * as React from 'react';
import {
  IconButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import {
  Attach24Regular,
  Delete24Regular,
  Document24Regular,
  ArrowDownload24Regular,
  DocumentPdf24Regular,
  DocumentText24Regular,
  DocumentTableRegular,
  DocumentBulletList24Regular,
  Folder24Regular,
  Image24Regular
} from '@fluentui/react-icons';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SharePointAlertService } from '../Services/SharePointAlertService';
import { logger } from '../Services/LoggerService';
import { NotificationService } from '../Services/NotificationService';
import { useFluentDialogs } from '../Hooks/useFluentDialogs';
import styles from './AttachmentManager.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

export interface IAttachment {
  fileName: string;
  serverRelativeUrl: string;
  size?: number;
}

export interface IAttachmentManagerProps {
  context: ApplicationCustomizerContext;
  alertService: SharePointAlertService;
  listId: string;
  itemId?: number;
  siteId?: string;
  attachments: IAttachment[];
  onAttachmentsChange: (attachments: IAttachment[]) => void;
  disabled?: boolean;
  maxFileSize?: number;
  allowedFormats?: string[];
}

interface IUploadProgress {
  fileName: string;
  progress: number;
  status: 'uploading' | 'completed' | 'error';
}

const UploadProgressFill: React.FC<IUploadProgress> = ({ progress, status }) => {
  const clampedProgress = Math.round(Math.max(0, Math.min(100, progress)));
  const progressClassMap = styles as unknown as Record<string, string>;
  const widthClassName =
    progressClassMap[`progressFill${clampedProgress}`] ||
    progressClassMap["progressFill0"] ||
    "";

  const statusClassName =
    status === "error"
      ? styles.progressFillError
      : status === "completed"
        ? styles.progressFillComplete
        : styles.progressFillUploading;

  return (
    <div
      className={`${styles.progressFill} ${widthClassName} ${statusClassName}`}
    />
  );
};

const AttachmentManager: React.FC<IAttachmentManagerProps> = ({
  context,
  alertService,
  listId,
  itemId,
  siteId,
  attachments,
  onAttachmentsChange,
  disabled = false,
  maxFileSize = 10,
  allowedFormats = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.txt', '.zip', '.jpg', '.jpeg', '.png', '.gif']
}) => {
  const [isUploading, setIsUploading] = React.useState(false);
  const [uploadProgress, setUploadProgress] = React.useState<IUploadProgress[]>([]);
  const [isDragOver, setIsDragOver] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const dropZoneRef = React.useRef<HTMLDivElement>(null);
  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );
  const { confirm, dialogs } = useFluentDialogs();

  const handleButtonClick = React.useCallback(() => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  }, []);

  const handleDragEnter = React.useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled && !isUploading && itemId) {
      setIsDragOver(true);
    }
  }, [disabled, isUploading, itemId]);

  const handleDragLeave = React.useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.currentTarget === dropZoneRef.current) {
      setIsDragOver(false);
    }
  }, []);

  const handleDragOver = React.useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDrop = React.useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);

    if (disabled || isUploading || !itemId) {
      return;
    }

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processFiles(Array.from(files));
    }
  }, [disabled, isUploading, itemId]);

  const uploadAttachment = React.useCallback(async (file: File, progressCallback: (progress: number) => void): Promise<IAttachment> => {
    if (!itemId) {
      throw new Error('Item ID is required to upload attachments');
    }

    try {
      logger.info('AttachmentManager', 'Starting attachment upload', { fileName: file.name, itemId });

      progressCallback(10);

      const arrayBuffer = await file.arrayBuffer();
      progressCallback(30);

      const result = await alertService.addAttachment(listId, itemId, file.name, arrayBuffer, siteId);

      progressCallback(90);

      const attachment: IAttachment = {
        fileName: result.fileName,
        serverRelativeUrl: result.serverRelativeUrl,
        size: file.size
      };

      progressCallback(100);
      logger.info('AttachmentManager', 'Attachment uploaded successfully', { fileName: attachment.fileName });

      return attachment;
    } catch (error) {
      logger.error('AttachmentManager', 'Failed to upload attachment', error);
      throw error;
    }
  }, [context, listId, itemId, siteId]);

  const processFiles = React.useCallback(async (filesToProcess: File[]) => {
    if (!itemId) {
      notificationService.showWarning(strings.AttachmentManagerSaveAlertFirst, strings.AttachmentManagerTitle);
      return;
    }

    const validFiles: File[] = [];
    const errors: string[] = [];

    filesToProcess.forEach(file => {
      const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
      const fileSizeMB = file.size / (1024 * 1024);

      if (!allowedFormats.includes(fileExtension)) {
        errors.push(Text.format(strings.AttachmentManagerInvalidFormat, file.name, allowedFormats.join(', ')));
      } else if (fileSizeMB > maxFileSize) {
        errors.push(Text.format(strings.AttachmentManagerFileTooLarge, file.name, maxFileSize.toString(), fileSizeMB.toFixed(2)));
      } else {
        validFiles.push(file);
      }
    });

    if (errors.length > 0) {
      notificationService.showWarning(errors.join(" "), strings.AttachmentManagerTitle);
    }

    if (validFiles.length === 0) {
      return;
    }

    setIsUploading(true);
    const newAttachments: IAttachment[] = [];

    const initialProgress: IUploadProgress[] = validFiles.map(file => ({
      fileName: file.name,
      progress: 0,
      status: 'uploading' as const
    }));
    setUploadProgress(initialProgress);

    try {
      for (let i = 0; i < validFiles.length; i++) {
        const file = validFiles[i];

        try {
          const attachment = await uploadAttachment(file, (progress) => {
            setUploadProgress(prev => prev.map((p, idx) =>
              idx === i ? { ...p, progress } : p
            ));
          });

          newAttachments.push(attachment);

          setUploadProgress(prev => prev.map((p, idx) =>
            idx === i ? { ...p, status: 'completed' as const, progress: 100 } : p
          ));
        } catch (error) {
          setUploadProgress(prev => prev.map((p, idx) =>
            idx === i ? { ...p, status: 'error' as const } : p
          ));
          logger.error('AttachmentManager', 'Failed to upload file', { fileName: file.name, error });
        }
      }

      if (newAttachments.length > 0) {
        onAttachmentsChange([...attachments, ...newAttachments]);
        logger.info('AttachmentManager', 'All attachments uploaded successfully', { count: newAttachments.length });
      }
    } catch (error) {
      logger.error('AttachmentManager', 'Attachment upload process failed', error);
      notificationService.showError(strings.AttachmentManagerUploadFailed, strings.AttachmentManagerTitle);
    } finally {
      setIsUploading(false);
      setTimeout(() => {
        setUploadProgress([]);
      }, 2000);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  }, [
    allowedFormats,
    maxFileSize,
    uploadAttachment,
    attachments,
    onAttachmentsChange,
    itemId,
    notificationService,
  ]);

  const handleFileChange = React.useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    await processFiles(Array.from(files));
  }, [processFiles]);

  const handleRemoveAttachment = React.useCallback(async (attachment: IAttachment) => {
    if (!itemId) return;

    const shouldDelete = await confirm({
      title: strings.AttachmentManagerTitle,
      message: Text.format(strings.AttachmentManagerDeleteConfirm, attachment.fileName),
      confirmText: strings.Delete,
    });
    if (!shouldDelete) {
      return;
    }

    try {
      await alertService.deleteAttachment(listId, itemId, attachment.fileName, siteId);

      const updatedAttachments = attachments.filter(a => a.fileName !== attachment.fileName);
      onAttachmentsChange(updatedAttachments);

      logger.info('AttachmentManager', 'Attachment deleted successfully', { fileName: attachment.fileName });
    } catch (error) {
      logger.error('AttachmentManager', 'Failed to delete attachment', error);
      notificationService.showError(strings.AttachmentManagerDeleteError, strings.AttachmentManagerTitle);
    }
  }, [context, listId, itemId, siteId, attachments, onAttachmentsChange, confirm, notificationService]);

  const formatFileSize = (bytes?: number): string => {
    if (!bytes) return '';

    const kb = bytes / 1024;
    if (kb < 1024) {
      return Text.format(strings.FileSizeKilobytes, kb.toFixed(1));
    }

    const mb = kb / 1024;
    return Text.format(strings.FileSizeMegabytes, mb.toFixed(1));
  };

  const getFileIcon = (fileName: string): React.ReactElement => {
    const extension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();

    const iconMap: { [key: string]: React.ReactElement } = {
      '.pdf': <DocumentPdf24Regular />,
      '.doc': <DocumentText24Regular />,
      '.docx': <DocumentText24Regular />,
      '.txt': <DocumentText24Regular />,
      '.xls': <DocumentTableRegular />,
      '.xlsx': <DocumentTableRegular />,
      '.csv': <DocumentTableRegular />,
      '.ppt': <DocumentBulletList24Regular />,
      '.pptx': <DocumentBulletList24Regular />,
      '.zip': <Folder24Regular />,
      '.rar': <Folder24Regular />,
      '.7z': <Folder24Regular />,
      '.jpg': <Image24Regular />,
      '.jpeg': <Image24Regular />,
      '.png': <Image24Regular />,
      '.gif': <Image24Regular />,
      '.bmp': <Image24Regular />,
      '.svg': <Image24Regular />
    };

    return iconMap[extension] || <Document24Regular />;
  };

  return (
    <div className={styles.attachmentManager}>
      {dialogs}
      <div className={styles.header}>
        <div className={styles.title}>{strings.AttachmentManagerTitle}</div>
        <input
          ref={fileInputRef}
          type="file"
          accept={allowedFormats.join(',')}
          onChange={handleFileChange}
          className={styles.fileInput}
          disabled={disabled || isUploading || !itemId}
          multiple
        />
        <PrimaryButton
          onRenderIcon={() =>
            isUploading ? <Spinner size={SpinnerSize.xSmall} /> : <Attach24Regular />
          }
          onClick={handleButtonClick}
          disabled={disabled || isUploading || !itemId}
          className={styles.uploadButton}
        >
          {isUploading ? strings.AttachmentManagerUploadingLabel : strings.AttachmentManagerAddFilesButton}
        </PrimaryButton>
      </div>

      {!itemId && (
        <div className={styles.infoMessage}>
          💡 {strings.AttachmentManagerSaveNotice}
        </div>
      )}

      {itemId && !isUploading && (
        <div
          ref={dropZoneRef}
          className={`${styles.dropZone} ${isDragOver ? styles.dragOver : ''}`}
          onDragEnter={handleDragEnter}
          onDragLeave={handleDragLeave}
          onDragOver={handleDragOver}
          onDrop={handleDrop}
        >
          <Attach24Regular className={styles.dropZoneIcon} />
          <div className={styles.dropZoneText}>
            {strings.AttachmentManagerDropZoneInstruction}
          </div>
          <div className={styles.dropZoneHint}>
            {Text.format(strings.AttachmentManagerDropZoneHint, allowedFormats.join(', '), maxFileSize.toString())}
          </div>
        </div>
      )}

      {uploadProgress.length > 0 && (
        <div className={styles.progressContainer}>
          <div className={styles.progressTitle}>{strings.AttachmentManagerProgressTitle}</div>
          {uploadProgress.map((progress, index) => (
            <div key={index} className={styles.progressItem}>
              <div className={styles.progressHeader}>
                <span className={styles.progressFileName}>{progress.fileName}</span>
                <span className={styles.progressStatus}>
                  {progress.status === 'completed' && strings.AttachmentManagerStatusCompleted}
                  {progress.status === 'error' && strings.AttachmentManagerStatusFailed}
                  {progress.status === 'uploading' && `${progress.progress}%`}
                </span>
              </div>
              <div className={styles.progressBar}>
                <UploadProgressFill
                  fileName={progress.fileName}
                  progress={progress.progress}
                  status={progress.status}
                />
              </div>
            </div>
          ))}
        </div>
      )}

      {attachments.length > 0 && (
        <div className={styles.attachmentsList}>
          {attachments.map((attachment, index) => (
            <div key={index} className={styles.attachmentItem}>
              <div className={styles.fileIcon}>
                {getFileIcon(attachment.fileName)}
              </div>
              <div className={styles.fileInfo}>
                <div className={styles.fileName}>{attachment.fileName}</div>
                {attachment.size && (
                  <div className={styles.fileSize}>{formatFileSize(attachment.size)}</div>
                )}
              </div>
              <div className={styles.actions}>
                <IconButton
                  onRenderIcon={() => <ArrowDownload24Regular />}
                  href={`${window.location.origin}${attachment.serverRelativeUrl}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  title={strings.AttachmentManagerDownloadTitle}
                  className={styles.iconButton}
                />
                <IconButton
                  onRenderIcon={() => <Delete24Regular />}
                  onClick={() => handleRemoveAttachment(attachment)}
                  disabled={!itemId}
                  className={`${styles.iconButton} ${styles.deleteButton}`}
                  title={strings.AttachmentManagerDeleteTitle}
                />
              </div>
            </div>
          ))}
        </div>
      )}

      <div className={styles.helpText}>
        {Text.format(strings.AttachmentManagerHelpText, allowedFormats.join(', '), maxFileSize.toString())}
      </div>
    </div>
  );
};

export default AttachmentManager;
