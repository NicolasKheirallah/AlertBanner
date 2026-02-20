import * as React from "react";
import JoditEditor from "jodit-react";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import EmojiPicker from "./EmojiPicker";
import ImageUpload from "./ImageUpload";
import styles from "./SharePointRichTextEditor.module.scss";

export interface IRichTextStyleOptions {
  showBold?: boolean;
  showItalic?: boolean;
  showUnderline?: boolean;
  showAlign?: boolean;
  showList?: boolean;
  showLink?: boolean;
  showMore?: boolean;
  showStyles?: boolean;
  showStrikethrough?: boolean;
  showSubscript?: boolean;
  showSuperscript?: boolean;
  showFontName?: boolean;
  showFontSize?: boolean;
  showFontColor?: boolean;
  showBackgroundColor?: boolean;
}

export interface ISharePointRichTextEditorProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  context?: ApplicationCustomizerContext;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
  description?: string;
  className?: string;
  id?: string;
  styleOptions?: IRichTextStyleOptions;
  maxLength?: number;
  minHeight?: number;
  maxHeight?: number;
  allowHTML?: boolean;
  restrictedElements?: string[];
  ariaLabel?: string;
  ariaDescribedBy?: string;
  debounceMs?: number;
  imageFolderName?: string;
  disableImageUpload?: boolean;
}

const MAX_IMAGE_WIDTH_PX = 300;

const escapeHtmlAttribute = (value: string): string => {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
};

const SharePointRichTextEditor: React.FC<ISharePointRichTextEditorProps> = ({
  label,
  value,
  onChange,
  context,
  placeholder = "Enter your message...",
  required = false,
  disabled = false,
  error,
  description,
  className,
  id,
  styleOptions,
  maxLength = 10000,
  minHeight = 120,
  maxHeight = 400,
  allowHTML = true,
  restrictedElements = ["script", "iframe", "object", "embed"],
  ariaLabel,
  ariaDescribedBy,
  debounceMs = 300,
  imageFolderName,
  disableImageUpload = false,
}) => {
  const [internalValue, setInternalValue] = React.useState(value);
  const [characterCount, setCharacterCount] = React.useState(0);
  const [validationError, setValidationError] = React.useState<string>("");
  const debounceRef = React.useRef<number>();
  const editorRef = React.useRef<any>(null);

  const uniqueId = React.useMemo(
    () => id || `richtext-${Math.random().toString(36).substring(2, 11)}`,
    [id],
  );

  React.useEffect(() => {
    setInternalValue(value);
    updateCharacterCount(value);
  }, [value]);

  const updateCharacterCount = React.useCallback((text: string) => {
    const textContent = (text || "").replace(/<[^>]*>/g, "");
    setCharacterCount(textContent.length);
  }, []);

  const validateContent = React.useCallback(
    (text: string): string => {
      if (!text && required) {
        return "This field is required";
      }

      const textContent = text.replace(/<[^>]*>/g, "");
      if (textContent.length > maxLength) {
        return `Content exceeds maximum length of ${maxLength} characters`;
      }

      if (!allowHTML && /<[^>]*>/g.test(text)) {
        return "HTML content is not allowed";
      }

      if (restrictedElements.length > 0) {
        const restrictedPattern = new RegExp(
          `<(${restrictedElements.join("|")})[^>]*>`,
          "gi",
        );
        if (restrictedPattern.test(text)) {
          return `The following HTML elements are not allowed: ${restrictedElements.join(", ")}`;
        }
      }

      return "";
    },
    [required, maxLength, allowHTML, restrictedElements],
  );

  const handleEditorChange = React.useCallback(
    (text: string) => {
      setInternalValue(text);
      updateCharacterCount(text);

      const validationErr = validateContent(text);
      setValidationError(validationErr);

      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }

      debounceRef.current = window.setTimeout(() => {
        onChange(text);
      }, debounceMs);
    },
    [validateContent, onChange, debounceMs, updateCharacterCount],
  );

  React.useEffect(() => {
    return () => {
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
    };
  }, []);

  const handleInsertUploadedImage = React.useCallback(
    (imageUrl: string, file: File) => {
      const defaultAltText = file.name
        .replace(/\.[^/.]+$/, "")
        .replace(/[_\-]+/g, " ")
        .trim();

      const altAttribute = escapeHtmlAttribute(defaultAltText || file.name);
      const imageMarkup = `<img src="${imageUrl}" alt="${altAttribute}" style="width:${MAX_IMAGE_WIDTH_PX}px;height:auto;" />`;

      if (editorRef.current) {
        try {
          // Jodit selection API
          editorRef.current.selection.insertHTML(imageMarkup);
        } catch (e) {
          // Fallback
          handleEditorChange((internalValue || "") + imageMarkup);
        }
      } else {
        handleEditorChange((internalValue || "") + imageMarkup);
      }
    },
    [internalValue, handleEditorChange],
  );

  const handleInsertEmoji = React.useCallback(
    (emoji: string) => {
      if (editorRef.current) {
        try {
          editorRef.current.selection.insertHTML(emoji);
        } catch (e) {
          handleEditorChange((internalValue || "") + emoji);
        }
      } else {
        handleEditorChange((internalValue || "") + emoji);
      }
    },
    [internalValue, handleEditorChange],
  );

  const currentError = error || validationError;
  const isOverLimit = characterCount > maxLength;
  const describedBy =
    ariaDescribedBy || (description ? `${uniqueId}-description` : undefined);

  // Configure Jodit
  const joditConfig = React.useMemo(
    () => ({
      readonly: disabled,
      placeholder: placeholder,
      minHeight: minHeight,
      maxHeight: maxHeight,
      toolbarAdaptive: false,
      controls: {
        customImagePicker: {
          icon: '<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2zm0 2v12h16V6H4zm3 8h10l-3.5-4.5-2.5 3-2-2.5L7 14z" /></svg>',
          tooltip: "Insert Image (SharePoint/Local)",
          exec: () => {
            document.getElementById(`${uniqueId}-image-btn`)?.click();
          },
        },
        customEmojiPicker: {
          icon: '<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z" /></svg>',
          tooltip: "Insert Emoji",
          exec: () => {
            document.getElementById(`${uniqueId}-emoji-btn`)?.click();
          },
        },
      },
      buttons: [
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "|",
        "ul",
        "ol",
        "|",
        "outdent",
        "indent",
        "|",
        "font",
        "fontsize",
        "brush",
        "paragraph",
        "|",
        "customImagePicker",
        "customEmojiPicker",
        "link",
        "|",
        "align",
        "undo",
        "redo",
        "hr",
        "eraser",
        "fullsize",
      ],
      // Disable built-in image upload since we use our custom uploader hook
      uploader: {
        insertImageAsBase64URI: false,
        url: "",
      },
      showCharsCounter: false,
      showWordsCounter: false,
      showXPathInStatusbar: false,
      askBeforePasteHTML: false,
      askBeforePasteFromWord: false,
      defaultActionOnPaste: "insert_as_html",
    }),
    [disabled, placeholder, minHeight, maxHeight],
  );

  return (
    <div
      className={`${styles.field} ${className || ""} ${currentError ? styles.error : ""}`}
    >
      <div className={styles.labelContainer}>
        <label
          className={styles.label}
          htmlFor={uniqueId}
          id={`${uniqueId}-label`}
        >
          {label}
          {required && <span className={styles.required}>*</span>}
        </label>

        {maxLength > 0 && (
          <div
            className={`${styles.characterCount} ${isOverLimit ? styles.overLimit : ""}`}
          >
            {characterCount} / {maxLength}
          </div>
        )}
      </div>

      {description && (
        <div className={styles.description} id={`${uniqueId}-description`}>
          {description}
        </div>
      )}

      {context && (
        <div className={styles.mediaToolbar} style={{ display: "none" }}>
          <EmojiPicker
            id={`${uniqueId}-emoji-btn`}
            onEmojiSelect={handleInsertEmoji}
            disabled={disabled}
          />
          {!disableImageUpload && (
            <ImageUpload
              id={`${uniqueId}-image-btn`}
              context={context}
              folderName={imageFolderName}
              onImageUploaded={handleInsertUploadedImage}
              disabled={disabled}
            />
          )}
        </div>
      )}

      <div
        className={`${styles.editorContainer} ${currentError ? styles.editorError : ""}`}
      >
        <JoditEditor
          ref={editorRef}
          value={internalValue}
          config={joditConfig as any}
          onBlur={(newContent) => handleEditorChange(newContent)}
          onChange={(newContent) => {
            updateCharacterCount(newContent);
            setValidationError(validateContent(newContent));
          }}
        />
      </div>

      {currentError && (
        <div
          className={styles.error}
          id={`${uniqueId}-error`}
          role="alert"
          aria-live="polite"
        >
          {currentError}
        </div>
      )}

      {/* Hidden helper text for screen readers */}
      <div className={styles.srOnly}>
        Rich text editor. Use toolbar buttons or keyboard shortcuts to format
        text.
        {maxLength > 0 && ` Maximum ${maxLength} characters allowed.`}
        {required && " This field is required."}
      </div>
    </div>
  );
};

export default SharePointRichTextEditor;
