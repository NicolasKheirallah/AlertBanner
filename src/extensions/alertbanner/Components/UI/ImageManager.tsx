import * as React from 'react';
import { Delete24Regular, Image24Regular, Copy24Regular } from '@fluentui/react-icons';
import { SharePointButton } from './SharePointControls';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { logger } from '../Services/LoggerService';
import { DateUtils } from '../Utils/DateUtils';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import styles from './ImageManager.module.scss';

export interface IImageManagerProps {
  context: ApplicationCustomizerContext;
  folderName: string; // Alert-specific folder name (e.g., languageGroup or title)
  onImageDeleted?: () => void;
}

interface IImageFile {
  name: string;
  serverRelativeUrl: string;
  timeCreated: string;
  length: number;
}

const ImageManager: React.FC<IImageManagerProps> = ({
  context,
  folderName,
  onImageDeleted
}) => {
  const [images, setImages] = React.useState<IImageFile[]>([]);

  const { loading: isLoading, error, execute: loadImages } = useAsyncOperation(
    async () => {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const serverRelativeUrl = context.pageContext.web.serverRelativeUrl;
      const cleanServerRelativeUrl = serverRelativeUrl.startsWith('/')
        ? serverRelativeUrl.substring(1)
        : serverRelativeUrl;

      const folderPath = cleanServerRelativeUrl
        ? `/${cleanServerRelativeUrl}/SiteAssets/AlertBannerImages/${folderName}`
        : `/SiteAssets/AlertBannerImages/${folderName}`;

      // Get files from folder
      const filesUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files?$select=Name,ServerRelativeUrl,TimeCreated,Length`;

      const response: SPHttpClientResponse = await context.spHttpClient.get(
        filesUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        if (response.status === 404) {
          // Folder doesn't exist yet - no images
          return [];
        }
        throw new Error(`Failed to load images: ${response.statusText}`);
      }

      const data = await response.json();
      const imageFiles: IImageFile[] = data.value
        .filter((file: any) => /\.(jpg|jpeg|png|gif|webp)$/i.test(file.Name))
        .map((file: any) => ({
          name: file.Name,
          serverRelativeUrl: file.ServerRelativeUrl,
          timeCreated: file.TimeCreated,
          length: file.Length
        }));

      logger.info('ImageManager', 'Loaded images', { count: imageFiles.length, folder: folderName });
      return imageFiles;
    },
    {
      onSuccess: (imageFiles) => setImages(imageFiles || []),
      logErrors: true
    }
  );

  React.useEffect(() => {
    loadImages();
  }, [loadImages]);

  const { execute: deleteImage } = useAsyncOperation(
    async (image: IImageFile) => {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const deleteUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(image.serverRelativeUrl)}')`;

      const response = await context.spHttpClient.post(
        deleteUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to delete image: ${response.statusText}`);
      }

      logger.info('ImageManager', 'Deleted image', { name: image.name });
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
        alert(`Failed to delete image: ${err.message}`);
      },
      logErrors: true
    }
  );

  const handleDeleteImage = React.useCallback(async (image: IImageFile) => {
    if (!confirm(`Delete "${image.name}"? This cannot be undone.`)) {
      return;
    }
    await deleteImage(image);
  }, [deleteImage]);

  const handleCopyUrl = React.useCallback((image: IImageFile) => {
    const fullUrl = `${window.location.origin}${image.serverRelativeUrl}`;
    navigator.clipboard.writeText(fullUrl).then(() => {
      alert('Image URL copied to clipboard!');
    }).catch(err => {
      logger.error('ImageManager', 'Error copying URL', err);
      alert('Failed to copy URL');
    });
  }, []);

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  if (isLoading) {
    return (
      <div className={styles.container}>
        <div className={styles.loading}>Loading images...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.container}>
        <div className={styles.error}>Error: {error.message}</div>
      </div>
    );
  }

  if (images.length === 0) {
    return (
      <div className={styles.container}>
        <div className={styles.emptyState}>
          <Image24Regular className={styles.emptyIcon} />
          <p>No images uploaded yet</p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h4>Uploaded Images ({images.length})</h4>
        <SharePointButton variant="secondary" onClick={loadImages}>
          Refresh
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
                Copy URL
              </SharePointButton>
              <SharePointButton
                variant="danger"
                icon={<Delete24Regular />}
                onClick={() => handleDeleteImage(image)}
              >
                Delete
              </SharePointButton>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default ImageManager;
