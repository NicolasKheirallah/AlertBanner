import * as React from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import styles from "./SharePointControls.module.scss";

const generateId = (() => {
  let counter = 0;
  return () => `sp-control-${counter++}`;
})();

export interface ISharePointButtonProps extends Omit<
  React.ButtonHTMLAttributes<HTMLButtonElement>,
  "onClick" | "disabled" | "type" | "className"
> {
  children?: React.ReactNode;
  onClick?: () => void;
  variant?: "primary" | "secondary" | "danger";
  disabled?: boolean;
  icon?: React.ReactNode;
  className?: string;
  type?: "button" | "submit" | "reset";
}

export const SharePointButton: React.FC<ISharePointButtonProps> = ({
  children,
  onClick,
  variant = "secondary",
  disabled = false,
  icon,
  className,
  type = "button",
  ...rest
}) => {
  const buttonClass = `${styles.button} ${styles[variant]} ${className || ""}`;

  return (
    <button
      type={type}
      className={buttonClass}
      onClick={onClick}
      disabled={disabled}
      {...rest}
    >
      {icon && <span className={styles.buttonIcon}>{icon}</span>}
      {children}
    </button>
  );
};

export interface ISharePointInputProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
  description?: string;
  type?:
    | "text"
    | "email"
    | "url"
    | "password"
    | "datetime-local"
    | "date"
    | "time";
  className?: string;
}

export const SharePointInput: React.FC<ISharePointInputProps> = ({
  label,
  value,
  onChange,
  placeholder,
  required = false,
  disabled = false,
  error,
  description,
  type = "text",
  className,
}) => {
  const inputId = React.useMemo(() => generateId(), []);
  const errorId = `${inputId}-error`;
  const descId = `${inputId}-desc`;

  return (
    <div className={`${styles.field} ${className || ""}`}>
      <label htmlFor={inputId} className={styles.label}>
        {label}
        {required && <span className={styles.required}>*</span>}
      </label>

      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}

      <input
        id={inputId}
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        disabled={disabled}
        className={`${styles.input} ${error ? styles.inputError : ""}`}
        aria-describedby={description ? descId : error ? errorId : undefined}
        aria-invalid={!!error}
        required={required}
      />

      {error && (
        <div id={errorId} className={styles.error}>
          {error}
        </div>
      )}
    </div>
  );
};

export interface ISharePointTextAreaProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
  description?: string;
  rows?: number;
  className?: string;
}

export const SharePointTextArea: React.FC<ISharePointTextAreaProps> = ({
  label,
  value,
  onChange,
  placeholder,
  required = false,
  disabled = false,
  error,
  description,
  rows = 4,
  className,
}) => {
  const textareaId = React.useMemo(() => generateId(), []);
  const errorId = `${textareaId}-error`;
  const descId = `${textareaId}-desc`;

  return (
    <div className={`${styles.field} ${className || ""}`}>
      <label htmlFor={textareaId} className={styles.label}>
        {label}
        {required && <span className={styles.required}>*</span>}
      </label>

      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}

      <textarea
        id={textareaId}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        disabled={disabled}
        rows={rows}
        className={`${styles.textarea} ${error ? styles.inputError : ""}`}
        aria-describedby={description ? descId : error ? errorId : undefined}
        aria-invalid={!!error}
        required={required}
      />

      {error && (
        <div id={errorId} className={styles.error}>
          {error}
        </div>
      )}
    </div>
  );
};

export interface ISharePointSelectOption {
  value: string;
  label: string;
  disabled?: boolean;
}

export interface ISharePointSelectProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: ISharePointSelectOption[];
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
  description?: string;
  className?: string;
}

export const SharePointSelect: React.FC<ISharePointSelectProps> = ({
  label,
  value,
  onChange,
  options,
  placeholder,
  required = false,
  disabled = false,
  error,
  description,
  className,
}) => {
  const selectId = React.useMemo(() => generateId(), []);
  const errorId = `${selectId}-error`;
  const descId = `${selectId}-desc`;

  return (
    <div className={`${styles.field} ${className || ""}`}>
      <label htmlFor={selectId} className={styles.label}>
        {label}
        {required && <span className={styles.required}>*</span>}
      </label>

      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}

      <select
        id={selectId}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
        className={`${styles.select} ${error ? styles.inputError : ""}`}
        aria-describedby={description ? descId : error ? errorId : undefined}
        aria-invalid={!!error}
        required={required}
      >
        {placeholder && (
          <option value="" disabled>
            {placeholder}
          </option>
        )}
        {options.map((option) => (
          <option
            key={option.value}
            value={option.value}
            disabled={option.disabled}
          >
            {option.label}
          </option>
        ))}
      </select>

      {error && (
        <div id={errorId} className={styles.error}>
          {error}
        </div>
      )}
    </div>
  );
};

export interface ISharePointToggleProps {
  label: string;
  checked: boolean;
  onChange: (checked: boolean) => void;
  disabled?: boolean;
  description?: string;
  className?: string;
}

export const SharePointToggle: React.FC<ISharePointToggleProps> = ({
  label,
  checked,
  onChange,
  disabled = false,
  description,
  className,
}) => {
  const toggleId = React.useMemo(() => generateId(), []);
  const descId = `${toggleId}-desc`;

  return (
    <div className={`${styles.field} ${className || ""}`}>
      <div className={styles.toggleContainer}>
        <button
          id={toggleId}
          type="button"
          role="switch"
          aria-checked={checked}
          aria-describedby={description ? descId : undefined}
          disabled={disabled}
          className={`${styles.toggle} ${checked ? styles.toggleOn : styles.toggleOff}`}
          onClick={() => onChange(!checked)}
        >
          <span className={styles.toggleThumb} />
        </button>

        <label htmlFor={toggleId} className={styles.toggleLabel}>
          {label}
        </label>
      </div>

      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}
    </div>
  );
};

export interface ISharePointSectionProps {
  title: string;
  children: React.ReactNode;
  collapsed?: boolean;
  onToggle?: (collapsed: boolean) => void;
  className?: string;
}

export const SharePointSection: React.FC<ISharePointSectionProps> = ({
  title,
  children,
  collapsed = false,
  onToggle,
  className,
}) => {
  return (
    <div className={`${styles.section} ${className || ""}`}>
      <div
        className={styles.sectionHeader}
        onClick={onToggle ? () => onToggle(!collapsed) : undefined}
        role={onToggle ? "button" : undefined}
        tabIndex={onToggle ? 0 : undefined}
      >
        <h3 className={styles.sectionTitle}>{title}</h3>
        {onToggle && (
          <span
            className={`${styles.sectionToggle} ${collapsed ? styles.collapsed : ""}`}
          >
            ⌄
          </span>
        )}
      </div>

      {!collapsed && <div className={styles.sectionContent}>{children}</div>}
    </div>
  );
};

export interface ISharePointPeoplePickerProps {
  context: any;
  titleText: string;
  personSelectionLimit?: number;
  groupName?: string; // Filter by SharePoint group
  showtooltip?: boolean;
  required?: boolean;
  disabled?: boolean;
  onChange?: (items: any[]) => void;
  defaultSelectedUsers?: string[];
  principalTypes?: PrincipalType[];
  resolveDelay?: number;
  className?: string;
  description?: string;
  error?: string;
}

export const SharePointPeoplePicker: React.FC<ISharePointPeoplePickerProps> = ({
  context,
  titleText,
  personSelectionLimit = 50,
  groupName,
  showtooltip = false,
  required = false,
  disabled = false,
  onChange,
  defaultSelectedUsers,
  principalTypes = [
    PrincipalType.User,
    PrincipalType.SharePointGroup,
    PrincipalType.SecurityGroup,
  ],
  resolveDelay = 1000,
  className,
  description,
  error,
}) => {
  const pickerId = React.useMemo(() => generateId(), []);
  const errorId = `${pickerId}-error`;
  const descId = `${pickerId}-desc`;

  // Defensive check: ensure context is defined and has necessary properties
  const isContextValid =
    context &&
    context.pageContext &&
    context.pageContext.web &&
    context.pageContext.web.absoluteUrl;

  if (!isContextValid) {
    // If context is invalid, we can try to fall back or show a warning.
    // For PeoplePicker, we absolutely need the web URL to call the API.
    // We can try to use window.location.origin + context.pageContext.web.serverRelativeUrl if available?
    let fallbackUrl = "";
    if (
      context &&
      context.pageContext &&
      context.pageContext.web &&
      context.pageContext.web.serverRelativeUrl
    ) {
      fallbackUrl =
        window.location.origin + context.pageContext.web.serverRelativeUrl;
      // Hack: patch the context object if we can
      try {
        context.pageContext.web.absoluteUrl = fallbackUrl;
      } catch (e) {
        // Immutable?
      }
    }

    if (
      !fallbackUrl &&
      (!context ||
        !context.pageContext ||
        !context.pageContext.web ||
        !context.pageContext.web.absoluteUrl)
    ) {
      return (
        <div className={`${styles.field} ${className || ""}`}>
          <div className={styles.error}>
            People Picker Error: Context is not properly initialized (missing
            absoluteUrl). Please refresh the page.
          </div>
        </div>
      );
    }
  }

  return (
    <div className={`${styles.field} ${className || ""}`}>
      <div className={styles.peoplePickerContainer}>
        <PeoplePicker
          context={context}
          titleText={titleText}
          personSelectionLimit={personSelectionLimit}
          groupName={groupName}
          showtooltip={showtooltip}
          required={required}
          disabled={disabled}
          onChange={onChange}
          defaultSelectedUsers={defaultSelectedUsers}
          principalTypes={principalTypes}
          resolveDelay={resolveDelay}
          ensureUser={true}
        />
      </div>

      {description && (
        <div id={descId} className={styles.description}>
          {description}
        </div>
      )}

      {error && (
        <div id={errorId} className={styles.error}>
          {error}
        </div>
      )}
    </div>
  );
};
