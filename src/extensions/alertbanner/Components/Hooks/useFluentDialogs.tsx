import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Textarea,
  Label,
} from "@fluentui/react-components";

export interface IConfirmDialogOptions {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
}

export interface IPromptDialogOptions extends IConfirmDialogOptions {
  label?: string;
  placeholder?: string;
  defaultValue?: string;
  multiline?: boolean;
  required?: boolean;
}

interface IConfirmDialogState extends IConfirmDialogOptions {
  open: boolean;
}

interface IPromptDialogState extends IPromptDialogOptions {
  open: boolean;
  value: string;
}

interface IUseFluentDialogsResult {
  confirm: (options: IConfirmDialogOptions) => Promise<boolean>;
  prompt: (options: IPromptDialogOptions) => Promise<string | null>;
  dialogs: React.ReactNode;
}

export const useFluentDialogs = (): IUseFluentDialogsResult => {
  const [confirmState, setConfirmState] = React.useState<IConfirmDialogState | null>(null);
  const [promptState, setPromptState] = React.useState<IPromptDialogState | null>(null);

  const confirmResolverRef = React.useRef<((value: boolean) => void) | null>(null);
  const promptResolverRef = React.useRef<((value: string | null) => void) | null>(null);

  const confirm = React.useCallback((options: IConfirmDialogOptions): Promise<boolean> => {
    return new Promise<boolean>((resolve) => {
      confirmResolverRef.current = resolve;
      setConfirmState({
        open: true,
        title: options.title,
        message: options.message,
        confirmText: options.confirmText || "Confirm",
        cancelText: options.cancelText || "Cancel",
      });
    });
  }, []);

  const closeConfirm = React.useCallback((value: boolean) => {
    setConfirmState((prev) => (prev ? { ...prev, open: false } : null));
    if (confirmResolverRef.current) {
      confirmResolverRef.current(value);
      confirmResolverRef.current = null;
    }
    setTimeout(() => setConfirmState(null), 0);
  }, []);

  const prompt = React.useCallback((options: IPromptDialogOptions): Promise<string | null> => {
    return new Promise<string | null>((resolve) => {
      promptResolverRef.current = resolve;
      setPromptState({
        open: true,
        title: options.title,
        message: options.message,
        confirmText: options.confirmText || "Submit",
        cancelText: options.cancelText || "Cancel",
        label: options.label,
        placeholder: options.placeholder,
        multiline: options.multiline !== false,
        required: !!options.required,
        defaultValue: options.defaultValue || "",
        value: options.defaultValue || "",
      });
    });
  }, []);

  const closePrompt = React.useCallback((value: string | null) => {
    setPromptState((prev) => (prev ? { ...prev, open: false } : null));
    if (promptResolverRef.current) {
      promptResolverRef.current(value);
      promptResolverRef.current = null;
    }
    setTimeout(() => setPromptState(null), 0);
  }, []);

  const dialogs = (
    <>
      <Dialog
        open={!!confirmState?.open}
        onOpenChange={(_, data) => {
          if (!data.open) {
            closeConfirm(false);
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{confirmState?.title}</DialogTitle>
            <DialogContent>{confirmState?.message}</DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => closeConfirm(false)}>
                {confirmState?.cancelText || "Cancel"}
              </Button>
              <Button appearance="primary" onClick={() => closeConfirm(true)}>
                {confirmState?.confirmText || "Confirm"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog
        open={!!promptState?.open}
        onOpenChange={(_, data) => {
          if (!data.open) {
            closePrompt(null);
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{promptState?.title}</DialogTitle>
            <DialogContent>
              <p>{promptState?.message}</p>
              {promptState?.label && <Label>{promptState.label}</Label>}
              <Textarea
                value={promptState?.value || ""}
                placeholder={promptState?.placeholder}
                onChange={(_, data) =>
                  setPromptState((prev) =>
                    prev ? { ...prev, value: data.value } : prev,
                  )
                }
                resize="vertical"
              />
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => closePrompt(null)}>
                {promptState?.cancelText || "Cancel"}
              </Button>
              <Button
                appearance="primary"
                disabled={!!promptState?.required && !promptState.value.trim()}
                onClick={() => closePrompt(promptState?.value ?? "")}
              >
                {promptState?.confirmText || "Submit"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );

  return {
    confirm,
    prompt,
    dialogs,
  };
};

