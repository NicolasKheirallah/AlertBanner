import * as React from 'react';
import { Button, Spinner, ProgressBar } from '@fluentui/react-components';
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
import { SPHttpClient } from '@microsoft/sp-http';
import { logger } from '../Services/LoggerService';
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
  listId: string;
  itemId?: number;
  attachments: IAttachment[];
  onAttachmentsChange: (attachments: IAttachment[]) => void;
  disabled?: boolean;
  maxFileSize?: number; // in MB
  allowedFormats?: string[];
}

interface IUploadProgress {
  fileName: string;
  progress: number; // 0-100
  status: 'uploading' | 'completed' | 'error';
}

const AttachmentManager: React.FC<IAttachmentManagerProps> = ({
  context,
  listId,
  itemId,
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

  const handleButtonClick = React.useCallback(() => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  }, []);

  // Drag and drop handlers
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
    // Only set to false if leaving the drop zone entirely
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

      // Simulate progress for better UX (SharePoint doesn't provide real progress)
      progressCallback(10);

      const siteUrl = context.pageContext.web.absoluteUrl;
      const uploadUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;

      const arrayBuffer = await file.arrayBuffer();
      progressCallback(30);

      const response = await context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/octet-stream',
          },
          body: arrayBuffer
        }
      );

      progressCallback(70);

      if (!response.ok) {
        throw new Error(`Upload failed: ${response.statusText}`);
      }

      const result = await response.json();
      progressCallback(90);

      const attachment: IAttachment = {
        fileName: result.d.FileName,
        serverRelativeUrl: result.d.ServerRelativeUrl,
        size: file.size
      };

      progressCallback(100);
      logger.info('AttachmentManager', 'Attachment uploaded successfully', { fileName: attachment.fileName });

      return attachment;
    } catch (error) {
      logger.error('AttachmentManager', 'Failed to upload attachment', error);
      throw error;
    }
  }, [context, listId, itemId]);

  const processFiles = React.useCallback(async (filesToProcess: File[]) => {
    if (!itemId) {
      alert('Please save the alert first before adding attachments.');
      return;
    }

    const validFiles: File[] = [];
    const errors: string[] = [];

    // Validate all files
    filesToProcess.forEach(file => {
      const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
      const fileSizeMB = file.size / (1024 * 1024);

      if (!allowedFormats.includes(fileExtension)) {
        errors.push(`${file.name}: Invalid file format (allowed: ${allowedFormats.join(', ')})`);
      } else if (fileSizeMB > maxFileSize) {
        errors.push(`${file.name}: File size exceeds ${maxFileSize}MB limit (${fileSizeMB.toFixed(2)}MB)`);
      } else {
        validFiles.push(file);
      }
    });

    if (errors.length > 0) {
      alert(`Some files were rejected:\n\n${errors.join('\n')}`);
    }

    if (validFiles.length === 0) {
      return;
    }

    setIsUploading(true);
    const newAttachments: IAttachment[] = [];

    // Initialize progress tracking for each file
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

          // Mark as completed
          setUploadProgress(prev => prev.map((p, idx) =>
            idx === i ? { ...p, status: 'completed' as const, progress: 100 } : p
          ));
        } catch (error) {
          // Mark as error
          setUploadProgress(prev => prev.map((p, idx) =>
            idx === i ? { ...p, status: 'error' as const } : p
          ));
          logger.error('AttachmentManager', 'Failed to upload file', { fileName: file.name, error });
        }
      }

      // Update attachments list
      if (newAttachments.length > 0) {
        onAttachmentsChange([...attachments, ...newAttachments]);
        logger.info('AttachmentManager', 'All attachments uploaded successfully', { count: newAttachments.length });
      }
    } catch (error) {
      logger.error('AttachmentManager', 'Attachment upload process failed', error);
      alert('Failed to upload some attachments. Please try again.');
    } finally {
      setIsUploading(false);
      // Clear progress after 2 seconds
      setTimeout(() => {
        setUploadProgress([]);
      }, 2000);
      // Reset input
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  }, [allowedFormats, maxFileSize, uploadAttachment, attachments, onAttachmentsChange, itemId]);

  const handleFileChange = React.useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    await processFiles(Array.from(files));
  }, [processFiles]);

  const handleRemoveAttachment = React.useCallback(async (attachment: IAttachment) => {
    if (!itemId) return;

    if (!confirm(Text.format(strings.AttachmentManagerDeleteConfirm, attachment.fileName))) {
      return;
    }

    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const deleteUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(attachment.fileName)}')`;

      await context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
          }
        }
      );

      // Update attachments list
      const updatedAttachments = attachments.filter(a => a.fileName !== attachment.fileName);
      onAttachmentsChange(updatedAttachments);

      logger.info('AttachmentManager', 'Attachment deleted successfully', { fileName: attachment.fileName });
    } catch (error) {
      logger.error('AttachmentManager', 'Failed to delete attachment', error);
      alert(strings.AttachmentManagerDeleteError);
    }
  }, [context, listId, itemId, attachments, onAttachmentsChange]);

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

    // Map file extensions to specific icons
    const iconMap: { [key: string]: React.ReactElement } = {
      // PDF
      '.pdf': <DocumentPdf24Regular />,

      // Text/Word
      '.doc': <DocumentText24Regular />,
      '.docx': <DocumentText24Regular />,
      '.txt': <DocumentText24Regular />,

      // Excel
      '.xls': <DocumentTableRegular />,
      '.xlsx': <DocumentTableRegular />,
      '.csv': <DocumentTableRegular />,

      // PowerPoint
      '.ppt': <DocumentBulletList24Regular />,
      '.pptx': <DocumentBulletList24Regular />,

      // Archives
      '.zip': <Folder24Regular />,
      '.rar': <Folder24Regular />,
      '.7z': <Folder24Regular />,

      // Images
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
        <Button
          icon={isUploading ? <Spinner size="tiny" /> : <Attach24Regular />}
          onClick={handleButtonClick}
          disabled={disabled || isUploading || !itemId}
          appearance="primary"
          size="small"
          className={styles.uploadButton}
        >
          {isUploading ? strings.AttachmentManagerUploadingLabel : strings.AttachmentManagerAddFilesButton}
        </Button>
      </div>

      {!itemId && (
        <div className={styles.infoMessage}>
          💡 {strings.AttachmentManagerSaveNotice}
        </div>
      )}

      {/* Drag and Drop Zone */}
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

      {/* Upload Progress Indicators */}
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
              <ProgressBar
                value={progress.progress}
                max={100}
                className={styles.progressBar}
                color={progress.status === 'error' ? 'error' : progress.status === 'completed' ? 'success' : 'brand'}
              />
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
                <Button
                  icon={<ArrowDownload24Regular />}
                  appearance="subtle"
                  size="small"
                  as="a"
                  href={`${window.location.origin}${attachment.serverRelativeUrl}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  title={strings.AttachmentManagerDownloadTitle}
                />
                <Button
                  icon={<Delete24Regular />}
                  appearance="subtle"
                  size="small"
                  onClick={() => handleRemoveAttachment(attachment)}
                  disabled={!itemId}
                  className={styles.deleteButton}
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
