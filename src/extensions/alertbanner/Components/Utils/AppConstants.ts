/**
 * Application-wide constants for AlertBanner
 * Consolidates magic strings and configuration values
 */

/**
 * SharePoint list names
 */
export const LIST_NAMES = {
  ALERTS: 'Alerts',
  ALERT_TYPES: 'AlertBannerTypes'
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
  TITLE_EN: 'Title_EN',
  TITLE_FR: 'Title_FR',
  TITLE_DE: 'Title_DE',
  TITLE_ES: 'Title_ES',
  TITLE_SV: 'Title_SV',
  TITLE_FI: 'Title_FI',
  TITLE_DA: 'Title_DA',
  TITLE_NO: 'Title_NO',

  DESCRIPTION_EN: 'Description_EN',
  DESCRIPTION_FR: 'Description_FR',
  DESCRIPTION_DE: 'Description_DE',
  DESCRIPTION_ES: 'Description_ES',
  DESCRIPTION_SV: 'Description_SV',
  DESCRIPTION_FI: 'Description_FI',
  DESCRIPTION_DA: 'Description_DA',
  DESCRIPTION_NO: 'Description_NO',

  LINK_DESCRIPTION_EN: 'LinkDescription_EN',
  LINK_DESCRIPTION_FR: 'LinkDescription_FR',
  LINK_DESCRIPTION_DE: 'LinkDescription_DE',
  LINK_DESCRIPTION_ES: 'LinkDescription_ES',
  LINK_DESCRIPTION_SV: 'LinkDescription_SV',
  LINK_DESCRIPTION_FI: 'LinkDescription_FI',
  LINK_DESCRIPTION_DA: 'LinkDescription_DA',
  LINK_DESCRIPTION_NO: 'LinkDescription_NO',

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
  // Allowed HTML tags
  ALLOWED_TAGS: [
    'p', 'br', 'strong', 'em', 'u', 'ul', 'ol', 'li',
    'a', 'span', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'blockquote', 'code', 'pre', 'img'
  ],

  // Allowed attributes
  ALLOWED_ATTRIBUTES: {
    'a': ['href', 'title', 'target'],
    'img': ['src', 'alt', 'title', 'width', 'height'],
    '*': ['class', 'id', 'style']
  },

  // Allowed URL schemes
  ALLOWED_URL_SCHEMES: ['http', 'https', 'mailto'],

  // Dangerous property names for JSON validation
  DANGEROUS_JSON_KEYS: ['__proto__', 'constructor', 'prototype']
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
 * Language codes mapping
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
