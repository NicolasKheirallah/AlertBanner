import * as React from "react";
import {
  TextBold24Regular,
  TextItalic24Regular,
  TextUnderline24Regular,
  TextBulletList24Regular,
  TextNumberListLtr24Regular,
  Link24Regular,
  TextAlignLeft24Regular,
  TextAlignCenter24Regular,
  TextAlignRight24Regular,
  Code24Regular,
  TextQuote24Regular
} from "@fluentui/react-icons";
import styles from "./SharePointRichTextEditor.module.scss";

export interface ISharePointRichTextEditorProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
  description?: string;
  className?: string;
  rows?: number;
}

const SharePointRichTextEditor: React.FC<ISharePointRichTextEditorProps> = ({
  label,
  value,
  onChange,
  placeholder = "Enter your message...",
  required = false,
  disabled = false,
  error,
  description,
  className,
  rows = 6
}) => {
  const editorRef = React.useRef<HTMLDivElement>(null);
  const [isToolbarVisible, setIsToolbarVisible] = React.useState(false);
  const [selectedRange, setSelectedRange] = React.useState<Range | null>(null);

  const editorId = React.useMemo(() => 
    `rich-editor-${Math.random().toString(36).substr(2, 9)}`, []
  );
  const errorId = `${editorId}-error`;
  const descId = `${editorId}-desc`;

  // Initialize editor content
  React.useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value) {
      editorRef.current.innerHTML = value || '';
    }
  }, [value]);

  const handleInput = React.useCallback(() => {
    if (editorRef.current) {
      const content = editorRef.current.innerHTML;
      onChange(content);
    }
  }, [onChange]);

  const handleFocus = () => {
    setIsToolbarVisible(true);
  };

  const handleBlur = (e: React.FocusEvent) => {
    // Only hide toolbar if we're not focusing on a toolbar button
    const relatedTarget = e.relatedTarget as HTMLElement;
    if (!relatedTarget || !relatedTarget.closest(`.${styles.toolbar}`)) {
      setTimeout(() => setIsToolbarVisible(false), 100);
    }
  };

  const saveSelection = () => {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      setSelectedRange(selection.getRangeAt(0));
    }
  };

  const restoreSelection = () => {
    if (selectedRange) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(selectedRange);
      }
    }
  };

  const executeCommand = (command: string, value?: string) => {
    restoreSelection();
    document.execCommand(command, false, value);
    handleInput();
    editorRef.current?.focus();
  };

  const insertHTML = (html: string) => {
    restoreSelection();
    document.execCommand('insertHTML', false, html);
    handleInput();
    editorRef.current?.focus();
  };

  const toggleFormat = (command: string) => {
    executeCommand(command);
  };

  const insertLink = () => {
    restoreSelection();
    const url = prompt('Enter URL:');
    if (url) {
      const selection = window.getSelection();
      const text = selection?.toString() || url;
      insertHTML(`<a href="${url}" target="_blank" rel="noopener noreferrer">${text}</a>`);
    }
  };

  const insertList = (ordered: boolean) => {
    executeCommand(ordered ? 'insertOrderedList' : 'insertUnorderedList');
  };

  const setAlignment = (align: string) => {
    executeCommand(`justify${align}`);
  };

  const insertQuote = () => {
    const selection = window.getSelection();
    const text = selection?.toString() || 'Quote text here...';
    insertHTML(`<blockquote style="border-left: 4px solid #0078d4; padding-left: 16px; margin: 16px 0; font-style: italic; color: #605e5c;">${text}</blockquote>`);
  };

  const insertCodeBlock = () => {
    const selection = window.getSelection();
    const text = selection?.toString() || 'Code here...';
    insertHTML(`<pre style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; padding: 12px; font-family: 'Courier New', monospace; font-size: 13px; overflow-x: auto;"><code>${text}</code></pre>`);
  };

  const toolbarButtons = [
    { 
      icon: <TextBold24Regular />, 
      command: 'bold', 
      title: 'Bold (Ctrl+B)',
      onClick: () => toggleFormat('bold')
    },
    { 
      icon: <TextItalic24Regular />, 
      command: 'italic', 
      title: 'Italic (Ctrl+I)',
      onClick: () => toggleFormat('italic')
    },
    { 
      icon: <TextUnderline24Regular />, 
      command: 'underline', 
      title: 'Underline (Ctrl+U)',
      onClick: () => toggleFormat('underline')
    },
    { 
      icon: <TextBulletList24Regular />, 
      title: 'Bullet List',
      onClick: () => insertList(false)
    },
    { 
      icon: <TextNumberListLtr24Regular />, 
      title: 'Numbered List',
      onClick: () => insertList(true)
    },
    { 
      icon: <Link24Regular />, 
      title: 'Insert Link',
      onClick: insertLink
    },
    { 
      icon: <TextAlignLeft24Regular />, 
      title: 'Align Left',
      onClick: () => setAlignment('Left')
    },
    { 
      icon: <TextAlignCenter24Regular />, 
      title: 'Align Center',
      onClick: () => setAlignment('Center')
    },
    { 
      icon: <TextAlignRight24Regular />, 
      title: 'Align Right',
      onClick: () => setAlignment('Right')
    },
    { 
      icon: <TextQuote24Regular />, 
      title: 'Insert Quote',
      onClick: insertQuote
    },
    { 
      icon: <Code24Regular />, 
      title: 'Insert Code Block',
      onClick: insertCodeBlock
    }
  ];

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'b':
          e.preventDefault();
          toggleFormat('bold');
          break;
        case 'i':
          e.preventDefault();
          toggleFormat('italic');
          break;
        case 'u':
          e.preventDefault();
          toggleFormat('underline');
          break;
        case 'k':
          e.preventDefault();
          insertLink();
          break;
      }
    }
  };

  return (
    <div className={`${styles.field} ${className || ''}`}>
      <label htmlFor={editorId} className={styles.label}>
        {label}
        {required && <span className={styles.required}>*</span>}
      </label>
      
      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}

      <div className={styles.editorContainer}>
        {isToolbarVisible && (
          <div className={styles.toolbar}>
            <div className={styles.toolbarGroup}>
              {toolbarButtons.map((button, index) => (
                <button
                  key={index}
                  type="button"
                  className={styles.toolbarButton}
                  title={button.title}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={(e) => {
                    e.preventDefault();
                    saveSelection();
                    button.onClick();
                  }}
                  disabled={disabled}
                >
                  {button.icon}
                </button>
              ))}
            </div>
          </div>
        )}
        
        <div
          ref={editorRef}
          id={editorId}
          className={`${styles.editor} ${error ? styles.editorError : ''}`}
          contentEditable={!disabled}
          suppressContentEditableWarning={true}
          onInput={handleInput}
          onFocus={handleFocus}
          onBlur={handleBlur}
          onMouseUp={saveSelection}
          onKeyUp={saveSelection}
          onKeyDown={handleKeyDown}
          style={{ minHeight: `${rows * 20}px` }}
          data-placeholder={placeholder}
          aria-describedby={description ? descId : error ? errorId : undefined}
          aria-invalid={!!error}
          aria-required={required}
          role="textbox"
          aria-multiline="true"
        />
      </div>
      
      {error && (
        <div id={errorId} className={styles.error}>
          {error}
        </div>
      )}
      
      <div className={styles.helpText}>
        <span>Rich text formatting available. Use Ctrl+B for bold, Ctrl+I for italic, Ctrl+U for underline, Ctrl+K for links.</span>
      </div>
    </div>
  );
};

export default SharePointRichTextEditor;