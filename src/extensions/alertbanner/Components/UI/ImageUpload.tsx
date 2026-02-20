import * as React from "react";
import { IconButton, TooltipHost } from "@fluentui/react";
import { ImageAdd24Regular } from "@fluentui/react-icons";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { ImageStorageService } from "../Services/ImageStorageService";
import { logger } from "../Services/LoggerService";
import { NotificationService } from "../Services/NotificationService";
import { useAsyncOperation } from "../Hooks/useAsyncOperation";
import styles from "./ImageUpload.module.scss";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import {
  FilePicker,
  IFilePickerResult,
} from "@pnp/spfx-controls-react/lib/FilePicker";

const ImageUpload: React.FC<{
  id?: string;
  context: ApplicationCustomizerContext;
  onImageUploaded: (
    imageUrl: string,
    file: File,
    widthPercent?: number,
  ) => void;
  folderName?: string;
  disabled?: boolean;
}> = ({ id, context, onImageUploaded, folderName, disabled = false }) => {
  const [isFilePickerOpen, setIsFilePickerOpen] = React.useState(false);
  const storageServiceRef = React.useRef<ImageStorageService>();

  if (!storageServiceRef.current) {
    storageServiceRef.current = new ImageStorageService(context);
  }
  const notificationService = React.useMemo(
    () => NotificationService.getInstance(context),
    [context],
  );

  const { loading: isUploading, execute: uploadImage } = useAsyncOperation(
    async (file: File) => {
      logger.info("ImageUpload", "Uploading local image", {
        name: file.name,
        size: file.size,
        type: file.type,
      });
      const imageUrl = await storageServiceRef.current!.uploadImage(
        file,
        folderName,
      );
      logger.info("ImageUpload", "Local image upload completed", {
        url: imageUrl,
      });
      return { imageUrl, file };
    },
    {
      onSuccess: (result) => {
        if (result) {
          onImageUploaded(result.imageUrl, result.file, 100);
        }
      },
      onError: (error) => {
        logger.error("ImageUpload", "Local image upload failed", error);
        const errorMessage =
          error instanceof Error ? error.message : strings.ImageUploadFailure;
        notificationService.showError(errorMessage, strings.ImageUploadFailure);
      },
      logErrors: true,
    },
  );

  const handleButtonClick = React.useCallback(() => {
    if (disabled || isUploading) {
      return;
    }
    setIsFilePickerOpen(true);
  }, [disabled, isUploading]);

  const handleFileSelected = React.useCallback(
    async (filePickerResult: IFilePickerResult[]) => {
      setIsFilePickerOpen(false);

      if (!filePickerResult || filePickerResult.length === 0) {
        return;
      }

      const selectedFile = filePickerResult[0];

      // Case 1: The user selected a file from SharePoint/OneDrive/Web Search
      if (selectedFile.fileAbsoluteUrl) {
        // We create a mock 'File' object just to satisfy the onImageUploaded signature, which Jodit uses for alt-text
        const mockFileName =
          selectedFile.fileName ||
          selectedFile.fileAbsoluteUrl.split("/").pop() ||
          "sharepoint-image";
        const mockFile = new File([], mockFileName, { type: "image/png" });
        onImageUploaded(selectedFile.fileAbsoluteUrl, mockFile, 100);
        return;
      }

      // Case 2: The user selected a local file using the "Upload" tab in the FilePicker
      if (selectedFile.downloadFileContent) {
        try {
          const fileBlob = await selectedFile.downloadFileContent();
          if (!fileBlob.type?.startsWith("image/")) {
            notificationService.showWarning(
              strings.ImageUploadInvalidFile,
              strings.ImageUploadFailure,
            );
            return;
          }

          // Convert Blob to File
          const file = new File(
            [fileBlob],
            selectedFile.fileName || "uploaded-image",
            { type: fileBlob.type },
          );
          await uploadImage(file);
        } catch (error) {
          logger.error(
            "ImageUpload",
            "Failed to process local file from FilePicker",
            error,
          );
          notificationService.showError(
            strings.ImageUploadFailure,
            strings.ImageUploadFailure,
          );
        }
      }
    },
    [uploadImage, notificationService, onImageUploaded],
  );

  const label = strings.UploadImage;

  return (
    <div className={styles.imageUploadContainer}>
      <TooltipHost content={label}>
        <IconButton
          id={id}
          onRenderIcon={() => <ImageAdd24Regular />}
          onClick={handleButtonClick}
          disabled={disabled || isUploading}
          className={styles.uploadButton}
          title={label}
          ariaLabel={label}
        />
      </TooltipHost>

      {isFilePickerOpen && (
        <FilePicker
          bingAPIKey="<BING API KEY>"
          accepts={[
            ".gif",
            ".jpg",
            ".jpeg",
            ".bmp",
            ".dib",
            ".tif",
            ".tiff",
            ".ico",
            ".png",
            ".jxr",
            ".svg",
          ]}
          buttonClassName={styles.hiddenFilePickerBtn} // We hide the actual PnP button since we trigger it via state
          onSave={handleFileSelected}
          onChange={(filePickerResult: IFilePickerResult[]) => {
            // If the user selects a local file, we trigger the save immediately rather than waiting for another click
            if (
              filePickerResult &&
              filePickerResult.length > 0 &&
              typeof filePickerResult[0].downloadFileContent === "function"
            ) {
              handleFileSelected(filePickerResult);
            }
          }}
          onCancel={() => setIsFilePickerOpen(false)}
          context={context as any}
          hideWebSearchTab={true} // BING API key needed for Web Search
          hideLinkUploadTab={false}
        />
      )}
    </div>
  );
};

export default ImageUpload;
