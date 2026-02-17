import * as React from 'react';
import { IconButton, TooltipHost } from "@fluentui/react";
import { ImageAdd24Regular } from '@fluentui/react-icons';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { ImageStorageService } from '../Services/ImageStorageService';
import { logger } from '../Services/LoggerService';
import { NotificationService } from '../Services/NotificationService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import styles from './ImageUpload.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';

export interface IImageUploadProps {
  context: ApplicationCustomizerContext;
  onImageUploaded: (imageUrl: string, file: File, widthPercent?: number) => void;
  folderName?: string;
  disabled?: boolean;
}

const ImageUpload: React.FC<IImageUploadProps> = ({
  context,
  onImageUploaded,
  folderName,
  disabled = false
}) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const storageServiceRef = React.useRef<ImageStorageService>();

  if (!storageServiceRef.current) {
    storageServiceRef.current = new ImageStorageService(context);
  }
  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );

  // Upload image using useAsyncOperation
  const { loading: isUploading, execute: uploadImage } = useAsyncOperation(
    async (file: File) => {
      logger.info('ImageUpload', 'Uploading image', {
        name: file.name,
        size: file.size,
        type: file.type
      });
      const imageUrl = await storageServiceRef.current!.uploadImage(file, folderName);
      logger.info('ImageUpload', 'Image upload completed', { url: imageUrl });
      return { imageUrl, file };
    },
    {
      onSuccess: (result) => {
        if (result) {
          onImageUploaded(result.imageUrl, result.file, 100);
        }
      },
      onError: (error) => {
        logger.error('ImageUpload', 'Image upload failed', error);
        const errorMessage = error instanceof Error ? error.message : strings.ImageUploadFailure;
        notificationService.showError(errorMessage, strings.ImageUploadFailure);
      },
      logErrors: true
    }
  );

  const handleButtonClick = React.useCallback(() => {
    if (disabled || isUploading) {
      return;
    }
    fileInputRef.current?.click();
  }, [disabled, isUploading]);

  const resetFileInput = (): void => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFileSelected = React.useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    resetFileInput();

    if (!file) {
      return;
    }

    if (!file.type?.startsWith('image/')) {
      notificationService.showWarning(strings.ImageUploadInvalidFile, strings.ImageUploadFailure);
      return;
    }

    await uploadImage(file);
  }, [uploadImage, notificationService]);

  const label = strings.UploadImage;

  return (
    <div className={styles.imageUploadContainer}>
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className={styles.fileInput}
        onChange={handleFileSelected}
        tabIndex={-1}
        aria-hidden={true}
      />

      <TooltipHost content={label}>
        <IconButton
          onRenderIcon={() => <ImageAdd24Regular />}
          onClick={handleButtonClick}
          disabled={disabled || isUploading}
          className={styles.uploadButton}
          title={label}
          ariaLabel={label}
        />
      </TooltipHost>
    </div>
  );
};

export default ImageUpload;
