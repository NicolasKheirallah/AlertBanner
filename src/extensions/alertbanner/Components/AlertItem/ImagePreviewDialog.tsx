import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  Button,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";
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
    <Dialog open={isOpen} onOpenChange={(event, data) => !data.open && onClose()}>
      <DialogSurface className={styles.dialogSurface}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={onClose}
              />
            }
          >
            {imageAlt || "Image Preview"}
          </DialogTitle>
          <DialogContent className={styles.dialogContent}>
            <img
              src={imageUrl}
              alt={imageAlt || "Full size preview"}
              className={styles.fullSizeImage}
              onClick={(e) => e.stopPropagation()}
            />
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
