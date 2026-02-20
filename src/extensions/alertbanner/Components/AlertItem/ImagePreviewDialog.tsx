import * as React from "react";
import * as ReactDOM from "react-dom";
import { Icon } from "@fluentui/react/lib/Icon";
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
  // Handle escape key
  React.useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        onClose();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isOpen, onClose]);

  // Prevent background scroll
  React.useEffect(() => {
    if (isOpen) {
      document.body.style.overflow = "hidden";
    } else {
      document.body.style.overflow = "";
    }
    return () => {
      document.body.style.overflow = "";
    };
  }, [isOpen]);

  if (!isOpen || typeof document === "undefined") {
    return null;
  }

  const lightboxContent = (
    <div
      className={styles.lightboxOverlay}
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-label="Image Preview"
    >
      <button
        className={styles.closeButton}
        onClick={onClose}
        aria-label="Close"
        autoFocus
      >
        <Icon iconName="Cancel" />
      </button>

      <div className={styles.imageContent} onClick={(e) => e.stopPropagation()}>
        <img
          src={imageUrl}
          alt={imageAlt || "Full size preview"}
          className={styles.fullSizeImage}
        />
        {imageAlt && imageAlt !== "Image" && (
          <div className={styles.imageCaption}>{imageAlt}</div>
        )}
      </div>
    </div>
  );

  // Render into a portal attached to document.body to ensure it floats above absolutely everything
  return ReactDOM.createPortal(lightboxContent, document.body);
};
