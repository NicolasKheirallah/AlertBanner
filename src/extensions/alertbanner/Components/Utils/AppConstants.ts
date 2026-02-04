/**
 * Application-wide constants for AlertBanner
 * Consolidates magic strings and configuration values
 */

import { AlertPriority } from "../Alerts/IAlerts";

/**
 * SharePoint list names
 */
export const LIST_NAMES = {
  ALERTS: 'Alerts',
  ALERT_TYPES: 'AlertBannerTypes'
} as const;

export const ALERT_ITEM_TYPES = {
  ALERT: "alert",
  TEMPLATE: "template",
  DRAFT: "draft",
  SETTINGS: "settings"
} as const;

/**
 * SharePoint field names
 */
export const FIELD_NAMES = {
  // Core fields
  TITLE: 'Title',
  DESCRIPTION: 'Description',
  ALERT_TYPE: 'AlertType',
  PRIORITY: 'Priority',
  IS_PINNED: 'IsPinned',
  STATUS: 'Status',

  // Scheduling fields
  SCHEDULED_START: 'ScheduledStart',
  SCHEDULED_END: 'ScheduledEnd',

  // Targeting fields
  TARGET_SITES: 'TargetSites',
  TARGET_USERS: 'TargetUsers',
  NOTIFICATION_TYPE: 'NotificationType',

  // Link fields
  LINK_URL: 'LinkUrl',
  LINK_DESCRIPTION: 'LinkDescription',

  // Language fields
  ITEM_TYPE: 'ItemType',
  TARGET_LANGUAGE: 'TargetLanguage',
  LANGUAGE_GROUP: 'LanguageGroup',
  AVAILABLE_FOR_ALL: 'AvailableForAll',

  // Multi-language content fields
  // Multi-language content fields
  // Note: We use Row-Based Localization, so we do NOT use Title_EN, Title_FR columns.
  // These are removed to avoid confusion.

  // Metadata fields
  METADATA: 'Metadata',
  CREATED: 'Created',
  CREATED_BY: 'CreatedBy',
  CREATED_DATE_TIME: 'CreatedDateTime',
  MODIFIED: 'Modified',
  AUTHOR: 'Author',
  ATTACHMENT_FILES: 'AttachmentFiles',
  ATTACHMENTS: 'Attachments'
} as const;

/**
 * Validation limits
 */
export const VALIDATION_LIMITS = {
  // Text length limits
  TITLE_MIN_LENGTH: 1,
  TITLE_MAX_LENGTH: 255,
  DESCRIPTION_MAX_LENGTH: 10000,
  LINK_URL_MAX_LENGTH: 2048,
  LINK_DESCRIPTION_MAX_LENGTH: 255,

  // Email validation
  EMAIL_MAX_LENGTH: 320,

  // URL validation
  URL_MAX_LENGTH: 2048,

  // JSON validation
  JSON_MAX_DEPTH: 10,

  // Date validation
  MAX_ALERT_DURATION_DAYS: 365,

  // Array limits
  MAX_TARGET_SITES: 100,
  MAX_TARGET_USERS: 100,

  // File size limits
  MAX_FILE_SIZE_MB: 10,
  MAX_FILE_SIZE_BYTES: 10 * 1024 * 1024
} as const;

/**
 * Cache configuration
 */
export const CACHE_CONFIG = {
  // Cache durations in milliseconds
  ALERTS_CACHE_DURATION: 5 * 60 * 1000, // 5 minutes
  USER_CACHE_DURATION: 15 * 60 * 1000, // 15 minutes
  SITE_CACHE_DURATION: 30 * 60 * 1000, // 30 minutes
  DEFAULT_CACHE_DURATION: 24 * 60 * 60 * 1000, // 24 hours

  // Storage keys
  STORAGE_PREFIX: 'alertbanner_',
  DISMISSED_ALERTS_KEY: 'alertbanner_dismissed_alerts',
  HIDDEN_ALERTS_KEY: 'alertbanner_hidden_alerts',
  USER_PREFERENCES_KEY: 'alertbanner_user_preferences',
  LANGUAGE_PREFERENCE_KEY: 'alertbanner_language_preference'
} as const;

/**
 * API configuration
 */
export const API_CONFIG = {
  // Graph API endpoints
  GRAPH_BASE_URL: 'https://graph.microsoft.com/v1.0',

  // Retry configuration
  MAX_RETRY_ATTEMPTS: 3,
  RETRY_DELAY_MS: 1000,
  RETRY_DELAY_MULTIPLIER: 2,

  // Request timeouts
  DEFAULT_TIMEOUT_MS: 30000, // 30 seconds
  LONG_TIMEOUT_MS: 60000, // 1 minute

  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_PAGE_SIZE: 100
} as const;

/**
 * HTML sanitization configuration
 */
export const SANITIZATION_CONFIG = {
  ALLOWED_TAGS: [
    'div', 'span', 'p', 'br', 'strong', 'b', 'em', 'i', 'u', 's', 'strike',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'ul', 'ol', 'li',
    'a', 'img',
    'blockquote', 'pre', 'code',
    'table', 'thead', 'tbody', 'tr', 'td', 'th'
  ],
  ALLOWED_ATTR: [
    'href', 'src', 'alt', 'title', 'class', 'id', 'style',
    'target', 'rel', 'width', 'height'
  ],
  FORBID_TAGS: ['script', 'object', 'embed', 'iframe', 'form', 'input'],
  FORBID_ATTR: ['onerror', 'onload', 'onclick', 'onmouseover'],
  DANGEROUS_JSON_KEYS: ['__proto__', 'constructor', 'prototype'],
  TRUSTED_DOMAINS: [
    /(^|\.)sharepoint\.[a-z.]+$/i,
    /(^|\.)office\.[a-z.]+$/i,
    /(^|\.)microsoft\.[a-z.]+$/i,
    /(^|\.)cdn\.jsdelivr\.net$/i,
    /(^|\.)cdnjs\.cloudflare\.com$/i
  ]
} as const;

export const VALIDATION_MESSAGES = {
  CreateAlertLanguageRequired: 'At least one language must be configured',
  CreateAlertLanguageTitleRequired: 'Title is required for {0}',
  CreateAlertLanguageDescriptionRequired: 'Description is required for {0}',
  CreateAlertLanguageLinkDescriptionRequired: 'Link description is required for {0} when URL is provided',
  CreateAlertLanguageAtLeastOneComplete: 'At least one language must have complete content (title and description)',
  CreateAlertDefaultLanguageRequired: 'Default language ({0}) must have complete content',
  DuplicateLanguagesNotAllowed: 'Duplicate language variants detected: {0}',
  TitleRequired: 'Title is required',
  TitleMinLength: 'Title must be at least 3 characters',
  TitleMaxLength: 'Title cannot exceed 100 characters',
  DescriptionRequired: 'Description is required',
  DescriptionMinLength: 'Description must be at least 10 characters',
  LinkDescriptionRequired: 'Link description is required when URL is provided',
  AlertTypeRequired: 'Alert type is required',
  InvalidUrlFormat: 'Please enter a valid URL',
  AtLeastOneSiteRequired: 'At least one target site must be selected',
  EndDateMustBeAfterStartDate: 'End date must be after start date',
} as const;

/**
 * UI configuration
 */
export const UI_CONFIG = {
  // Animation durations
  FADE_IN_DURATION_MS: 300,
  FADE_OUT_DURATION_MS: 200,
  SLIDE_DURATION_MS: 250,

  // Auto-save intervals
  AUTO_SAVE_INTERVAL_MS: 30000, // 30 seconds
  AUTO_SAVE_DEBOUNCE_MS: 2000, // 2 seconds

  // Toast/notification durations
  TOAST_DURATION_SHORT_MS: 3000,
  TOAST_DURATION_MEDIUM_MS: 5000,
  TOAST_DURATION_LONG_MS: 8000,

  // Loading delays
  MIN_LOADING_DELAY_MS: 300,

  // Truncation lengths
  PREVIEW_DESCRIPTION_LENGTH: 150,
  CARD_TITLE_LENGTH: 60,
  LIST_ITEM_DESCRIPTION_LENGTH: 100
} as const;

/**
 * Alert type defaults
 */
export const ALERT_TYPE_DEFAULTS = {
  DEFAULT_TYPE: 'Info',
  DEFAULT_ICON: 'Info',
  DEFAULT_BACKGROUND_COLOR: '#0078d4',
  DEFAULT_TEXT_COLOR: '#ffffff',
  DEFAULT_LINK_TEXT: 'Learn More'
} as const;

/**
 * Default Alert Types Configuration
 */
// Default Alert Type Name when none is specified
export const DEFAULT_ALERT_TYPE_NAME = "Info";

export const DEFAULT_ALERT_TYPES = [
  {
    name: "Info",
    iconName: "Info",
    backgroundColor: "#389899",
    textColor: "#ffffff",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  },
  {
    name: "Warning",
    iconName: "Warning",
    backgroundColor: "#f1c40f",
    textColor: "#000000",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  },
  {
    name: "Maintenance",
    iconName: "ConstructionCone",
    backgroundColor: "#afd6d6",
    textColor: "#000000",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  },
  {
    name: "Interruption",
    iconName: "Error",
    backgroundColor: "#c54644",
    textColor: "#ffffff",
    additionalStyles: "",
    priorityStyles: {
      [AlertPriority.Critical]: "border: 2px solid #E81123;",
      [AlertPriority.High]: "border: 1px solid #EA4300;",
      [AlertPriority.Medium]: "",
      [AlertPriority.Low]: ""
    }
  }
] as const;

/**
 * Notification Styles
 */
export const NOTIFICATION_STYLES = {
  SUCCESS: { 
    backgroundColor: '#dff6dd', 
    textColor: '#107c10', 
    borderColor: '#107c10', 
    icon: '✅' 
  },
  WARNING: { 
    backgroundColor: '#fff4ce', 
    textColor: '#797673', 
    borderColor: '#ffb900', 
    icon: '⚠️' 
  },
  ERROR: { 
    backgroundColor: '#fde7e9', 
    textColor: '#a4262c', 
    borderColor: '#d13438', 
    icon: '❌' 
  },
  INFO: { 
    backgroundColor: '#deecf9', 
    textColor: '#323130', 
    borderColor: '#0078d4', 
    icon: 'ℹ️' 
  }
} as const;

/**
 * Status values
 */
export const ALERT_STATUS = {
  ACTIVE: 'Active',
  EXPIRED: 'Expired',
  SCHEDULED: 'Scheduled',
  DRAFT: 'Draft'
} as const;

/**
 * Auto-save prefixes
 */
export const AUTO_SAVE_PREFIXES = {
  LOWERCASE: '[auto-saved]',
  CAPITALIZED: '[Auto-saved]'
} as const;

/**
 * Supported languages configuration
 * Single source of truth for language definitions
 */
export const SUPPORTED_LANGUAGES = [
  { code: 'en-us', suffix: 'EN', name: 'English', nativeName: 'English', flag: '🇺🇸' },
  { code: 'fr-fr', suffix: 'FR', name: 'French', nativeName: 'Français', flag: '🇫🇷' },
  { code: 'de-de', suffix: 'DE', name: 'German', nativeName: 'Deutsch', flag: '🇩🇪' },
  { code: 'es-es', suffix: 'ES', name: 'Spanish', nativeName: 'Español', flag: '🇪🇸' },
  { code: 'sv-se', suffix: 'SV', name: 'Swedish', nativeName: 'Svenska', flag: '🇸🇪' },
  { code: 'fi-fi', suffix: 'FI', name: 'Finnish', nativeName: 'Suomi', flag: '🇫🇮' },
  { code: 'da-dk', suffix: 'DA', name: 'Danish', nativeName: 'Dansk', flag: '🇩🇰' },
  { code: 'nb-no', suffix: 'NO', name: 'Norwegian', nativeName: 'Norsk', flag: '🇳🇴' }
] as const;

/**
 * Language codes mapping (Derived/Legacy support)
 */
export const LANGUAGE_CODES = {
  'en-us': 'EN',
  'fr-fr': 'FR',
  'de-de': 'DE',
  'es-es': 'ES',
  'sv-se': 'SV',
  'fi-fi': 'FI',
  'da-dk': 'DA',
  'nb-no': 'NO'
} as const;

/**
 * Error messages (generic/fallback)
 */
export const ERROR_MESSAGES = {
  GENERIC_ERROR: 'An unexpected error occurred',
  NETWORK_ERROR: 'Network error occurred. Please check your connection',
  PERMISSION_DENIED: 'You do not have permission to perform this action',
  NOT_FOUND: 'The requested resource was not found',
  VALIDATION_FAILED: 'Validation failed. Please check your input',
  TIMEOUT: 'The request timed out. Please try again',
  UNAUTHORIZED: 'You are not authorized to access this resource'
} as const;

/**
 * Success messages (generic/fallback)
 */
export const SUCCESS_MESSAGES = {
  SAVED: 'Changes saved successfully',
  CREATED: 'Item created successfully',
  UPDATED: 'Item updated successfully',
  DELETED: 'Item deleted successfully',
  COPIED: 'Copied to clipboard'
} as const;

/**
 * Regular expressions for validation
 */
export const REGEX_PATTERNS = {
  EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
  URL: /^https?:\/\/.+/i,
  GUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
  LANGUAGE_FIELD: /^(Title|Description|LinkDescription)_[A-Z]{2}$/,
  HEX_COLOR: /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/,
  ALPHANUMERIC: /^[a-zA-Z0-9]+$/,
  ALPHANUMERIC_WITH_SPACES: /^[a-zA-Z0-9\s]+$/
} as const;

/**
 * HTTP status codes
 */
export const HTTP_STATUS = {
  OK: 200,
  CREATED: 201,
  NO_CONTENT: 204,
  BAD_REQUEST: 400,
  UNAUTHORIZED: 401,
  FORBIDDEN: 403,
  NOT_FOUND: 404,
  CONFLICT: 409,
  TOO_MANY_REQUESTS: 429,
  INTERNAL_SERVER_ERROR: 500,
  SERVICE_UNAVAILABLE: 503
} as const;

/**
 * GraphQL/OData operators
 */
export const ODATA_OPERATORS = {
  EQUAL: 'eq',
  NOT_EQUAL: 'ne',
  GREATER_THAN: 'gt',
  GREATER_THAN_OR_EQUAL: 'ge',
  LESS_THAN: 'lt',
  LESS_THAN_OR_EQUAL: 'le',
  AND: 'and',
  OR: 'or',
  NOT: 'not',
  NULL: 'null'
} as const;

/**
 * Feature flags (for gradual rollout)
 */
export const FEATURE_FLAGS = {
  ENABLE_MULTI_LANGUAGE: true,
  ENABLE_USER_TARGETING: true,
  ENABLE_NOTIFICATIONS: true,
  ENABLE_ATTACHMENTS: true,
  ENABLE_TEMPLATES: true,
  ENABLE_DRAFTS: true,
  ENABLE_AUTO_SAVE: true,
  ENABLE_RICH_TEXT_EDITOR: true,
  ENABLE_ADVANCED_SCHEDULING: true,
  ENABLE_SITE_TARGETING: true
} as const;

/**
 * UI Shadow configurations
 */
export const SHADOW_CONFIG = {
  CRITICAL_PRIORITY: '0 4px 12px rgba(232, 17, 35, 0.15)',
  HIGH_PRIORITY: '0 2px 8px rgba(234, 67, 0, 0.1)',
  MEDIUM_PRIORITY: '0 1px 4px rgba(0, 120, 212, 0.08)',
  CARD_HOVER: '0 2px 8px rgba(0, 0, 0, 0.1)',
  DIALOG: '0 8px 32px rgba(0, 0, 0, 0.15)'
} as const;

/**
 * Window.open configuration for security
 */
export const WINDOW_OPEN_CONFIG = {
  TARGET: '_blank',
  FEATURES: 'noopener,noreferrer'
} as const;

/**
 * UI Image configuration
 */
export const UI_IMAGE_CONFIG = {
  MAX_WIDTH_PX: 300,
  MIN_WIDTH_PX: 50,
  MAX_HEIGHT_PX: 400,
  DEFAULT_QUALITY: 0.8
} as const;

/**
 * Carousel configuration
 */
export const CAROUSEL_CONFIG = {
  MIN_INTERVAL: 2000,
  MAX_INTERVAL: 30000,
  DEFAULT_INTERVAL: 5000,
  MIN_SLIDES: 1,
  MAX_SLIDES: 10
} as const;

/**
 * Animation durations (in milliseconds)
 */
export const ANIMATION_DURATION = {
  FAST: 150,
  NORMAL: 300,
  SLOW: 500,
  CAROUSEL_TRANSITION: 400
} as const;

/**
 * Type guards for const assertions
 */
export type ListName = typeof LIST_NAMES[keyof typeof LIST_NAMES];
export type FieldName = typeof FIELD_NAMES[keyof typeof FIELD_NAMES];
export type AlertStatus = typeof ALERT_STATUS[keyof typeof ALERT_STATUS];
export type LanguageCode = keyof typeof LANGUAGE_CODES;
export type LanguageSuffix = typeof LANGUAGE_CODES[LanguageCode];
