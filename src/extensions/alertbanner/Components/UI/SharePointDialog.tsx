import * as React from "react";
import styles from "./SharePointDialog.module.scss";
import { Dismiss24Regular } from "@fluentui/react-icons";

export interface ISharePointDialogProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  width?: number | string;
  maxWidth?: number | string;
  height?: number | string;
  children: React.ReactNode;
  footer?: React.ReactNode;
  className?: string;
}

const SharePointDialog: React.FC<ISharePointDialogProps> = ({
  isOpen,
  onClose,
  title,
  width = 800,
  maxWidth,
  height,
  children,
  footer,
  className,
}) => {
  const dialogRef = React.useRef<HTMLDivElement>(null);

  const handleOverlayClick = (e: React.MouseEvent) => {
    if (e.target === e.currentTarget) {
      onClose();
    }
  };

  React.useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === "Escape" && isOpen) {
        onClose();
      }
    };

    if (isOpen) {
      document.addEventListener("keydown", handleEscape);
      document.body.style.overflow = "hidden";
    }

    return () => {
      document.removeEventListener("keydown", handleEscape);
      document.body.style.overflow = "unset";
    };
  }, [isOpen, onClose]);

  React.useEffect(() => {
    if (isOpen && dialogRef.current) {
      const focusableElement = dialogRef.current.querySelector(
        'button, input, select, textarea, [tabindex]:not([tabindex="-1"])',
      ) as HTMLElement;
      if (focusableElement) {
        focusableElement.focus();
      }
    }
  }, [isOpen]);

  if (!isOpen) {
    return null;
  }

  const dialogStyle: React.CSSProperties = {
    width,
    ...(maxWidth && { maxWidth }),
    ...(height && { height, maxHeight: height }),
  };

  return (
    <div className={styles.overlay} onClick={handleOverlayClick}>
      <div
        className={`${styles.dialog} ${className || ""}`}
        style={dialogStyle}
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby="dialog-title"
      >
        <div className={styles.header}>
          <h2 id="dialog-title" className={styles.title}>
            {title}
          </h2>
          <button
            className={styles.closeButton}
            onClick={onClose}
            aria-label="Close"
            type="button"
          >
            <Dismiss24Regular />
          </button>
        </div>

        <div className={styles.content}>{children}</div>

        {footer && <div className={styles.footer}>{footer}</div>}
      </div>
    </div>
  );
};

export default SharePointDialog;
