import * as React from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  PrimaryButton,
  TextField,
  Label,
} from "@fluentui/react";

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
        hidden={!confirmState?.open}
        onDismiss={() => closeConfirm(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: confirmState?.title,
          subText: confirmState?.message,
        }}
        modalProps={{
          isBlocking: false,
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => closeConfirm(false)}>
            {confirmState?.cancelText || "Cancel"}
          </DefaultButton>
          <PrimaryButton onClick={() => closeConfirm(true)}>
            {confirmState?.confirmText || "Confirm"}
          </PrimaryButton>
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!promptState?.open}
        onDismiss={() => closePrompt(null)}
        dialogContentProps={{
          type: DialogType.normal,
          title: promptState?.title,
        }}
        modalProps={{
          isBlocking: false,
        }}
      >
        <p>{promptState?.message}</p>
        {promptState?.label && <Label>{promptState.label}</Label>}
        <TextField
          multiline={promptState?.multiline !== false}
          rows={3}
          value={promptState?.value || ""}
          placeholder={promptState?.placeholder}
          onChange={(_, value) =>
            setPromptState((prev) => (prev ? { ...prev, value: value || "" } : prev))
          }
        />
        <DialogFooter>
          <DefaultButton onClick={() => closePrompt(null)}>
            {promptState?.cancelText || "Cancel"}
          </DefaultButton>
          <PrimaryButton
            disabled={!!promptState?.required && !promptState.value.trim()}
            onClick={() => closePrompt(promptState?.value ?? "")}
          >
            {promptState?.confirmText || "Submit"}
          </PrimaryButton>
        </DialogFooter>
      </Dialog>
    </>
  );

  return {
    confirm,
    prompt,
    dialogs,
  };
};
