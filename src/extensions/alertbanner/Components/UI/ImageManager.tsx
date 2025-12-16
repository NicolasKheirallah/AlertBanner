import * as React from 'react';
import { Delete24Regular, Image24Regular, Copy24Regular } from '@fluentui/react-icons';
import { SharePointButton } from './SharePointControls';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { ImageStorageService, IExistingImage } from '../Services/ImageStorageService';
import { logger } from '../Services/LoggerService';
import { DateUtils } from '../Utils/DateUtils';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import styles from './ImageManager.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

export interface IImageManagerProps {
  context: ApplicationCustomizerContext;
  imageStorageService: ImageStorageService;
  folderName: string; // Alert-specific folder name (e.g., languageGroup or title)
  onImageDeleted?: () => void;
}



const ImageManager: React.FC<IImageManagerProps> = ({
  context,
  imageStorageService,
  folderName,
  onImageDeleted
}) => {
  const [images, setImages] = React.useState<IExistingImage[]>([]);

  const fetchFolderImages = React.useCallback(async (): Promise<IExistingImage[]> => {
    return await imageStorageService.listImages(folderName);
  }, [imageStorageService, folderName]);

  const { loading: isLoading, error, execute: loadImages } = useAsyncOperation(
    fetchFolderImages,
    {
      onSuccess: (imageFiles) => setImages(imageFiles || []),
      logErrors: true
    }
  );

  const loadImagesRef = React.useRef(loadImages);
  React.useEffect(() => {
    loadImagesRef.current = loadImages;
  }, [loadImages]);

  React.useEffect(() => {
    loadImagesRef.current();
  }, [folderName, fetchFolderImages]);

  const { execute: deleteImage } = useAsyncOperation(
    async (image: IExistingImage) => {
      await imageStorageService.deleteImage(image.name, folderName);
      return true;
    },
    {
      onSuccess: async () => {
        await loadImages();
        if (onImageDeleted) {
          onImageDeleted();
        }
      },
      onError: (err) => {
        alert(Text.format(strings.ImageManagerDeleteError, err.message || ''));
      },
      logErrors: true
    }
  );

  const handleDeleteImage = React.useCallback(async (image: IExistingImage) => {
    if (!confirm(Text.format(strings.ImageManagerDeleteConfirm, image.name))) {
      return;
    }
    await deleteImage(image);
  }, [deleteImage]);

  const handleCopyUrl = React.useCallback((image: IExistingImage) => {
    const fullUrl = `${window.location.origin}${image.serverRelativeUrl}`;
    navigator.clipboard.writeText(fullUrl).then(() => {
      alert(strings.ImageManagerCopySuccess);
    }).catch(err => {
      logger.error('ImageManager', 'Error copying URL', err);
      alert(strings.ImageManagerCopyError);
    });
  }, []);

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return Text.format(strings.FileSizeKilobytes, (bytes / 1024).toFixed(1));
    return Text.format(strings.FileSizeMegabytes, (bytes / (1024 * 1024)).toFixed(1));
  };

  if (isLoading) {
    return (
      <div className={styles.container}>
        <div className={styles.loading}>{strings.ImageManagerLoading}</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.container}>
        <div className={styles.error}>{Text.format(strings.ImageManagerError, error.message || '')}</div>
      </div>
    );
  }

  if (images.length === 0) {
    return (
      <div className={styles.container}>
        <div className={styles.emptyState}>
          <Image24Regular className={styles.emptyIcon} />
          <p>{strings.ImageManagerEmpty}</p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h4>{Text.format(strings.ImageManagerHeaderTitle, images.length.toString())}</h4>
        <SharePointButton variant="secondary" onClick={loadImages}>
          {strings.Refresh}
        </SharePointButton>
      </div>

      <div className={styles.imageGrid}>
        {images.map(image => (
          <div key={image.serverRelativeUrl} className={styles.imageCard}>
            <div className={styles.imagePreview}>
              <img
                src={`${window.location.origin}${image.serverRelativeUrl}`}
                alt={image.name}
                onError={(e) => {
                  (e.target as HTMLImageElement).src = 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCIgZmlsbD0iI2YwZjBmMCIvPjx0ZXh0IHg9IjUwJSIgeT0iNTAlIiBmb250LWZhbWlseT0iQXJpYWwiIGZvbnQtc2l6ZT0iMTQiIGZpbGw9IiM5OTkiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGR5PSIuM2VtIj5JbWFnZSBub3QgZm91bmQ8L3RleHQ+PC9zdmc+';
                }}
              />
            </div>
            <div className={styles.imageInfo}>
              <div className={styles.imageName} title={image.name}>{image.name}</div>
              <div className={styles.imageMeta}>
                {formatFileSize(image.length)} • {DateUtils.formatForDisplay(image.timeCreated)}
              </div>
            </div>
            <div className={styles.imageActions}>
              <SharePointButton
                variant="secondary"
                icon={<Copy24Regular />}
                onClick={() => handleCopyUrl(image)}
              >
                {strings.ImageManagerCopyUrl}
              </SharePointButton>
              <SharePointButton
                variant="danger"
                icon={<Delete24Regular />}
                onClick={() => handleDeleteImage(image)}
              >
                {strings.Delete}
              </SharePointButton>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default ImageManager;
