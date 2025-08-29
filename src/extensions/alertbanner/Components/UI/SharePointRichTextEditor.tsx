import * as React from "react";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "./SharePointRichTextEditor.module.scss";

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
}

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
  className
}) => {

  const handleEditorChange = (text: string): string => {
    onChange(text);
    return text;
  };

  return (
    <div className={`${styles.field} ${className || ''} ${error ? styles.error : ''}`}>
      <label className={styles.label}>
        {label}
        {required && <span className={styles.required}>*</span>}
      </label>

      {description && (
        <div className={styles.description}>
          {description}
        </div>
      )}

      <div className={`${styles.editorContainer} ${error ? styles.editorError : ''}`}>
        <RichText
          value={value}
          onChange={handleEditorChange}
          placeholder={placeholder}
          className={styles.editor}
          isEditMode={!disabled}
        />
      </div>

      {error && (
        <div className={styles.error}>
          {error}
        </div>
      )}
    </div>
  );
};

export default SharePointRichTextEditor;
