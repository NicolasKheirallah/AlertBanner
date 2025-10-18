import * as React from 'react';
import { Button, Tooltip } from '@fluentui/react-components';
import { ImageAdd24Regular } from '@fluentui/react-icons';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { ImageStorageService } from '../Services/ImageStorageService';
import { logger } from '../Services/LoggerService';
import styles from './ImageUpload.module.scss';
import { useLocalizationContext } from '../Hooks/useLocalization';

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
  const [isUploading, setIsUploading] = React.useState(false);
  const storageServiceRef = React.useRef<ImageStorageService>();
  const { getString } = useLocalizationContext();

  if (!storageServiceRef.current) {
    storageServiceRef.current = new ImageStorageService(context);
  }

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
      alert('Please select a valid image file (PNG, JPG, GIF, or WebP).');
      return;
    }

    setIsUploading(true);

    try {
      logger.info('ImageUpload', 'Uploading image', {
        name: file.name,
        size: file.size,
        type: file.type
      });
      const imageUrl = await storageServiceRef.current!.uploadImage(file, folderName);
      logger.info('ImageUpload', 'Image upload completed', { url: imageUrl });
      onImageUploaded(imageUrl, file, 100);
    } catch (error) {
      logger.error('ImageUpload', 'Image upload failed', error);
      alert(error instanceof Error ? error.message : 'Image upload failed. Please try again.');
    } finally {
      setIsUploading(false);
    }
  }, [folderName, onImageUploaded]);

  const label = getString('UploadImage');

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

      <Tooltip content={label} relationship="label">
        <Button
          icon={<ImageAdd24Regular />}
          appearance="subtle"
          onClick={handleButtonClick}
          disabled={disabled || isUploading}
          className={styles.uploadButton}
          title={label}
          aria-label={label}
        />
      </Tooltip>
    </div>
  );
};

export default ImageUpload;
