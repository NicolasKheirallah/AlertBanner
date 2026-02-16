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

const BODY_SCROLL_LOCK_ATTR = "data-sharepoint-dialog-lock-count";
let dialogIdCounter = 0;

const toCssDimension = (value: number | string): string =>
  typeof value === "number" ? `${value}px` : value;

const lockBodyScroll = (): void => {
  const currentCount = Number(
    document.body.getAttribute(BODY_SCROLL_LOCK_ATTR) || "0",
  );
  const nextCount = currentCount + 1;
  document.body.setAttribute(BODY_SCROLL_LOCK_ATTR, String(nextCount));
  document.body.style.overflow = "hidden";
};

const unlockBodyScroll = (): void => {
  const currentCount = Number(
    document.body.getAttribute(BODY_SCROLL_LOCK_ATTR) || "0",
  );
  const nextCount = Math.max(0, currentCount - 1);

  if (nextCount === 0) {
    document.body.removeAttribute(BODY_SCROLL_LOCK_ATTR);
    document.body.style.overflow = "";
    return;
  }

  document.body.setAttribute(BODY_SCROLL_LOCK_ATTR, String(nextCount));
};

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
  const dialogTitleId = React.useMemo(() => {
    dialogIdCounter += 1;
    return `alert-dialog-title-${dialogIdCounter}`;
  }, []);

  const handleOverlayClick = (e: React.MouseEvent) => {
    if (e.target === e.currentTarget) {
      onClose();
    }
  };

  React.useEffect(() => {
    if (!isOpen) {
      return;
    }

    const handleEscape = (e: KeyboardEvent): void => {
      if (e.key === "Escape") {
        onClose();
      }
    };

    document.addEventListener("keydown", handleEscape);
    lockBodyScroll();

    return () => {
      document.removeEventListener("keydown", handleEscape);
      unlockBodyScroll();
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

  const handleDialogKeyDown = React.useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (event.key !== "Tab" || !dialogRef.current) {
        return;
      }

      const focusableElements = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])',
        ),
      ).filter((el) => {
        const style = window.getComputedStyle(el);
        return style.display !== "none" && style.visibility !== "hidden";
      });

      if (focusableElements.length === 0) {
        event.preventDefault();
        return;
      }

      const firstElement = focusableElements[0];
      const lastElement = focusableElements[focusableElements.length - 1];
      const activeElement = document.activeElement as HTMLElement | null;

      if (!event.shiftKey && activeElement === lastElement) {
        event.preventDefault();
        firstElement.focus();
      } else if (event.shiftKey && activeElement === firstElement) {
        event.preventDefault();
        lastElement.focus();
      }
    },
    [],
  );

  if (!isOpen) {
    return null;
  }

  const resolvedHeight = height ? toCssDimension(height) : undefined;
  const dialogStyle: React.CSSProperties = {
    width,
    ...(maxWidth && { maxWidth }),
    ...(resolvedHeight && {
      height: `min(${resolvedHeight}, calc(100vh - 40px))`,
      maxHeight: "calc(100vh - 40px)",
    }),
  };

  return (
    <div className={styles.overlay} onClick={handleOverlayClick}>
      <div
        className={`${styles.dialog} ${className || ""}`}
        style={dialogStyle}
        ref={dialogRef}
        onKeyDown={handleDialogKeyDown}
        role="dialog"
        aria-modal="true"
        aria-labelledby={dialogTitleId}
      >
        <div className={styles.header}>
          <h2 id={dialogTitleId} className={styles.title}>
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

        <div
          className={styles.notificationHost}
          data-alert-banner-dialog-notification-host="true"
        />

        <div className={styles.content}>{children}</div>

        {footer && <div className={styles.footer}>{footer}</div>}
      </div>
    </div>
  );
};

export default SharePointDialog;
