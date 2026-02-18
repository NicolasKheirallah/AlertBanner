import * as React from "react";
import {
  Dialog,
  DialogType,
} from "@fluentui/react";
import styles from "./ImagePreviewDialog.module.scss";

interface IImagePreviewDialogProps {
  isOpen: boolean;
  imageUrl: string;
  imageAlt?: string;
  onClose: () => void;
}

export const ImagePreviewDialog: React.FC<IImagePreviewDialogProps> = ({
  isOpen,
  imageUrl,
  imageAlt,
  onClose,
}) => {
  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onClose}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: imageAlt || "Image Preview",
        closeButtonAriaLabel: "Close",
      }}
      modalProps={{
        isBlocking: true,
        className: styles.imagePreviewModal,
        containerClassName: styles.imagePreviewContainer,
      }}
    >
      <div className={styles.dialogContent}>
        <img
          src={imageUrl}
          alt={imageAlt || "Full size preview"}
          className={styles.fullSizeImage}
          onClick={(e) => e.stopPropagation()}
          role="img"
        />
      </div>
    </Dialog>
  );
};
