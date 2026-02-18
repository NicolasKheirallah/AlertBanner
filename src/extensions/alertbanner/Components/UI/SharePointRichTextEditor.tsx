import * as React from "react";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import EmojiPicker from './EmojiPicker';
import ImageUpload from './ImageUpload';
import styles from "./SharePointRichTextEditor.module.scss";
import { createPortal } from "react-dom";

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
  // Enhanced PnP-specific options
  id?: string;
  styleOptions?: IRichTextStyleOptions;
  maxLength?: number;
  minHeight?: number;
  maxHeight?: number;
  // Content validation options
  allowHTML?: boolean;
  restrictedElements?: string[];
  // Accessibility enhancements
  ariaLabel?: string;
  ariaDescribedBy?: string;
  // Performance options
  debounceMs?: number;
  // Image upload folder customization
  imageFolderName?: string; // Custom folder name for image uploads (e.g., alert title)
  // Disable image upload for cross-site editing scenarios
  disableImageUpload?: boolean;
}

const escapeHtmlAttribute = (value: string): string => {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
};

const MAX_IMAGE_WIDTH_PX = 300;

let quillImageResizePromise: Promise<any> | null = null;
const ensureQuillImageResizeModule = (): Promise<any> => {
  if (typeof window === 'undefined') {
    return Promise.reject(new Error('Image resize module requires a browser environment.'));
  }

  if (!quillImageResizePromise) {
    quillImageResizePromise = (async () => {
      const globalAny = window as any;
      let Quill = globalAny.Quill;

      if (!Quill) {
        const quillModule: any = await import('quill');
        Quill = quillModule?.default || quillModule;
        if (!Quill) {
          throw new Error('Unable to load Quill editor');
        }
        globalAny.Quill = Quill;
      }

      const quillWithImports: any = Quill;
      if (!quillWithImports.imports || !quillWithImports.imports['modules/imageResize']) {
        const ImageResizeModule = (await import('quill-image-resize-module')).default;
        Quill.register('modules/imageResize', ImageResizeModule);
      }

      return Quill;
    })().catch(error => {
      quillImageResizePromise = null;
      throw error;
    });
  }

  return quillImageResizePromise;
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
  // Enhanced options
  id,
  styleOptions,
  maxLength = 10000,
  minHeight = 120,
  maxHeight = 400,
  allowHTML = true,
  restrictedElements = ['script', 'iframe', 'object', 'embed'],
  ariaLabel,
  ariaDescribedBy,
  debounceMs = 300,
  imageFolderName,
  disableImageUpload = false
}) => {
  const [internalValue, setInternalValue] = React.useState(value);
  const [characterCount, setCharacterCount] = React.useState(0);
  const [validationError, setValidationError] = React.useState<string>('');
  const debounceRef = React.useRef<number>();
  const richTextRef = React.useRef<RichText | null>(null);
  const uniqueId = React.useMemo(() => id || `richtext-${Math.random().toString(36).substring(2, 11)}`, [id]);
  const [toolbarElement, setToolbarElement] = React.useState<HTMLElement | null>(null);

  const clampImageElement = React.useCallback((image: HTMLImageElement): void => {
    const parsePxWidth = (value: string): number | undefined => {
      if (!value) {
        return undefined;
      }
      const match = value.match(/([\d.]+)px/i);
      if (!match) {
        return undefined;
      }
      const parsed = parseFloat(match[1]);
      return isNaN(parsed) ? undefined : parsed;
    };

    const inlineWidth = parsePxWidth(image.style.width);

    let targetWidth = inlineWidth;
    if (targetWidth === undefined || targetWidth <= 0) {
      const naturalWidth = image.naturalWidth || image.width || MAX_IMAGE_WIDTH_PX;
      targetWidth = naturalWidth;
    }

    targetWidth = Math.min(Math.max(targetWidth || MAX_IMAGE_WIDTH_PX, 1), MAX_IMAGE_WIDTH_PX);

    image.style.width = `${targetWidth}px`;
    image.style.height = 'auto';
    image.removeAttribute('width');
    image.removeAttribute('height');
  }, []);

  const clampAllImages = React.useCallback((root: HTMLElement | null) => {
    if (!root) {
      return;
    }

    const images = root.querySelectorAll('img');
    images.forEach(img => clampImageElement(img as HTMLImageElement));
  }, [clampImageElement]);

  // Update internal value when prop changes
  React.useEffect(() => {
    setInternalValue(value);
    updateCharacterCount(value);
  }, [value]);

  // Update character count
  const updateCharacterCount = React.useCallback((text: string) => {
    // Strip HTML tags for character counting
    const textContent = (text || '').replace(/<[^>]*>/g, '');
    setCharacterCount(textContent.length);
  }, []);

  // Enhanced validation
  const validateContent = React.useCallback((text: string): string => {
    if (!text && required) {
      return 'This field is required';
    }

    const textContent = text.replace(/<[^>]*>/g, '');
    if (textContent.length > maxLength) {
      return `Content exceeds maximum length of ${maxLength} characters`;
    }

    // Check for restricted elements
    if (!allowHTML && /<[^>]*>/g.test(text)) {
      return 'HTML content is not allowed';
    }

    if (restrictedElements.length > 0) {
      const restrictedPattern = new RegExp(`<(${restrictedElements.join('|')})[^>]*>`, 'gi');
      if (restrictedPattern.test(text)) {
        return `The following HTML elements are not allowed: ${restrictedElements.join(', ')}`;
      }
    }

    return '';
  }, [required, maxLength, allowHTML, restrictedElements]);

  // Enhanced change handler with debouncing and validation
  const handleEditorChange = React.useCallback((text: string): string => {
    // Update internal state immediately for responsive UI (with original text)
    setInternalValue(text);
    updateCharacterCount(text);

    // Validate content immediately for UI feedback
    const error = validateContent(text);
    setValidationError(error);

    // Clear previous debounce timeout
    if (debounceRef.current) {
      clearTimeout(debounceRef.current);
    }

    // Debounced onChange to parent (sanitization disabled to prevent encoding loops)
    debounceRef.current = window.setTimeout(() => {
      // SharePoint already provides security measures, so we skip aggressive sanitization
      // This prevents HTML encoding loops that cause text corruption
      onChange(text);
    }, debounceMs);

    return text; // Return original text immediately to avoid render issues
  }, [validateContent, onChange, debounceMs, updateCharacterCount]);

  // Cleanup on unmount
  React.useEffect(() => {
    return () => {
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
    };
  }, []);

  // Enhanced style options with security defaults
  const defaultStyleOptions: IRichTextStyleOptions = {
    showBold: true,
    showItalic: true,
    showUnderline: true,
    showAlign: true,
    showList: true,
    showLink: true,
    showMore: false, // Disable by default for security
    showStyles: true,
    showStrikethrough: false,
    showSubscript: false,
    showSuperscript: false,
    showFontName: false, // Disable to maintain consistent branding
    showFontSize: false, // Disable to maintain consistent sizing
    showFontColor: false, // Disable to maintain accessibility
    showBackgroundColor: false, // Disable to maintain accessibility
    ...styleOptions
  };

  const editorStyles: React.CSSProperties = {
    minHeight: `${minHeight}px`,
    maxHeight: `${maxHeight}px`
  };

  const currentError = error || validationError;
  const isOverLimit = characterCount > maxLength;
  const describedBy = ariaDescribedBy || (description ? `${uniqueId}-description` : undefined);

  const handleInsertUploadedImage = React.useCallback((imageUrl: string, file: File, _requestedWidth?: number) => {
    const editorInstance = richTextRef.current?.getEditor?.();
    const defaultAltText = file.name
      .replace(/\.[^/.]+$/, '')
      .replace(/[_\-]+/g, ' ')
      .trim();

    if (editorInstance) {
      const selection = editorInstance.getSelection(true);
      const insertIndex = selection ? selection.index : editorInstance.getLength();
      editorInstance.insertEmbed(insertIndex, 'image', imageUrl, 'user');
      editorInstance.setSelection(insertIndex + 1, 0);

      window.setTimeout(() => {
        const root = editorInstance.root as HTMLElement;
        const images = Array.from(root.querySelectorAll('img')) as HTMLImageElement[];
        const insertedImage = images.find(img => img.src === imageUrl) || images[images.length - 1];
        if (insertedImage) {
          insertedImage.alt = defaultAltText || file.name;
          const ensureClamp = () => clampImageElement(insertedImage);
          if (insertedImage.complete) {
            ensureClamp();
          } else {
            insertedImage.addEventListener('load', ensureClamp, { once: true });
          }
        }
        handleEditorChange(root.innerHTML);
      }, 0);

      return;
    }

    const altAttribute = escapeHtmlAttribute(defaultAltText || file.name);
    const imageMarkup = `<p><img src="${imageUrl}" alt="${altAttribute}" style="width:${MAX_IMAGE_WIDTH_PX}px;height:auto;" /></p>`;
    const newHtml = internalValue ? `${internalValue}${imageMarkup}` : imageMarkup;
    handleEditorChange(newHtml);
  }, [internalValue, handleEditorChange, clampImageElement]);

  React.useEffect(() => {
    let disposed = false;

    if (typeof window === 'undefined') {
      return undefined;
    }

    const attachModule = () => {
      const editor = richTextRef.current?.getEditor?.();
      if (!editor) {
        if (!disposed) {
          window.setTimeout(attachModule, 100);
        }
        return;
      }

      ensureQuillImageResizeModule()
        .then(Quill => {
          if (disposed) {
            return;
          }

          const editorAny = editor as any;
          const existing = editorAny.getModule?.('imageResize');
          if (existing) {
            return;
          }

          const options = {
            modules: ['Resize', 'DisplaySize'],
            displayStyles: {
              backgroundColor: 'rgba(255,255,255,0.85)',
              border: '1px solid #605e5c',
              color: '#323130',
              fontSize: '12px',
              padding: '2px 4px',
              borderRadius: '2px'
            }
          };

          const currentOptions = editorAny.options || {};
          const currentModules = currentOptions.modules || {};

          editorAny.options = {
            ...currentOptions,
            modules: {
              ...currentModules,
              imageResize: {
                ...(currentModules.imageResize || {}),
                ...options
              }
            }
          };

          const ModuleCtor = Quill.import('modules/imageResize');
          // Suppress module logging
          const originalLog = console.log;
          console.log = () => {};
          try {
            editorAny.modules = {
              ...(editorAny.modules || {}),
              imageResize: new ModuleCtor(editorAny, editorAny.options.modules.imageResize)
            };
          } finally {
            console.log = originalLog;
          }
        })
        .catch(error => {
          console.error('Failed to initialize Quill image resize module', error);
        });
    };

    attachModule();

    return () => {
      disposed = true;
    };
  }, []);

  React.useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    let cancelled = false;

    const locateToolbar = () => {
      const editor = richTextRef.current?.getEditor?.();
      const toolbar = editor?.getModule?.('toolbar')?.container as HTMLElement | undefined;

      if (!toolbar) {
        if (!cancelled) {
          window.setTimeout(locateToolbar, 100);
        }
        return;
      }

      if (!cancelled) {
        setToolbarElement(toolbar);
      }
    };

    locateToolbar();

    return () => {
      cancelled = true;
      setToolbarElement(null);
    };
  }, [uniqueId]);

  React.useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    let disposed = false;
    let cleanup: (() => void) | undefined;

    const attachClamp = () => {
      const editor = richTextRef.current?.getEditor?.();
      if (!editor) {
        if (!disposed) {
          window.setTimeout(attachClamp, 100);
        }
        return;
      }

      const editorAny = editor as any;
      const root = editor.root as HTMLElement;

      const applyClamp = () => clampAllImages(root);
      applyClamp();

      const handler = () => applyClamp();
      editorAny.on?.('text-change', handler);

      cleanup = () => {
        editorAny.off?.('text-change', handler);
      };
    };

    attachClamp();

    return () => {
      disposed = true;
      cleanup?.();
    };
  }, [clampAllImages]);

  const handleInsertEmoji = React.useCallback((emoji: string) => {
    const editorInstance = richTextRef.current?.getEditor?.();

    if (editorInstance) {
      const selection = editorInstance.getSelection(true);
      const insertIndex = selection ? selection.index : editorInstance.getLength();

      editorInstance.insertText(insertIndex, emoji, 'user');
      editorInstance.setSelection(insertIndex + emoji.length, 0);

      window.setTimeout(() => {
        const root = editorInstance.root as HTMLElement;
        handleEditorChange(root.innerHTML);
      }, 0);

      return;
    }

    const newValue = (internalValue || '') + emoji;
    handleEditorChange(newValue);
  }, [handleEditorChange, internalValue]);

  const renderMediaToolbar = React.useCallback(() => {
    if (!context) {
      return null;
    }

    return (
      <div className={styles.mediaToolbar}>
        <EmojiPicker
          onEmojiSelect={handleInsertEmoji}
          disabled={disabled}
        />
        {!disableImageUpload && (
          <ImageUpload
            context={context}
            folderName={imageFolderName}
            onImageUploaded={handleInsertUploadedImage}
            disabled={disabled}
          />
        )}
      </div>
    );
  }, [context, disabled, disableImageUpload, handleInsertEmoji, handleInsertUploadedImage, imageFolderName]);

  return (
    <div className={`${styles.field} ${className || ''} ${currentError ? styles.error : ''}`}>
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
          <div className={`${styles.characterCount} ${isOverLimit ? styles.overLimit : ''}`}>
            {characterCount} / {maxLength}
          </div>
        )}
      </div>

      {description && (
        <div
          className={styles.description}
          id={`${uniqueId}-description`}
        >
          {description}
        </div>
      )}

      <div className={`${styles.editorContainer} ${currentError ? styles.editorError : ''}`}>
        <RichText
          id={uniqueId}
          ref={richTextRef}
          value={internalValue}
          onChange={handleEditorChange}
          placeholder={placeholder}
          className={styles.editor}
          style={editorStyles}
          styleOptions={defaultStyleOptions}
          isEditMode={!disabled}
          // Accessibility enhancements
          {...(ariaLabel && { 'aria-label': ariaLabel })}
          {...(describedBy && { 'aria-describedby': describedBy })}
          aria-labelledby={`${uniqueId}-label`}
          aria-required={required}
          aria-invalid={!!currentError}
        />

        {context && toolbarElement && createPortal(renderMediaToolbar(), toolbarElement)}
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
        Rich text editor. Use toolbar buttons or keyboard shortcuts to format text.
        {maxLength > 0 && ` Maximum ${maxLength} characters allowed.`}
        {required && ' This field is required.'}
      </div>
    </div>
  );
};

export default SharePointRichTextEditor;
