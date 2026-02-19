import { AlertPriority, IAlertType } from "../Alerts/IAlerts";

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

  // Note: We use Row-Based Localization, so we do NOT use Title_EN, Title_FR columns.

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

  // File size limits
  MAX_FILE_SIZE_MB: 10,
  MAX_FILE_SIZE_BYTES: 10 * 1024 * 1024
} as const;

export const CACHE_CONFIG = {
  ALERTS_CACHE_DURATION: 5 * 60 * 1000, // 5 minutes

  // Storage keys
  STORAGE_PREFIX: 'alertbanner_',
  DISMISSED_ALERTS_KEY: 'alertbanner_dismissed_alerts',
  HIDDEN_ALERTS_KEY: 'alertbanner_hidden_alerts'
} as const;

export const API_CONFIG = {
  GRAPH_BASE_URL: 'https://graph.microsoft.com/v1.0',

  MAX_PAGE_SIZE: 100,

  MAX_RETRY_ATTEMPTS: 3,
  RETRY_DELAY_MS: 1000,
  RETRY_DELAY_MULTIPLIER: 2,

  DEFAULT_TIMEOUT_MS: 30000 // 30 seconds
} as const;

// Window open configuration for link actions
export const WINDOW_OPEN_CONFIG = {
  TARGET: '_blank',
  FEATURES: 'noopener,noreferrer'
} as const;

// Shadow configuration for critical priority alerts
export const SHADOW_CONFIG = {
  CRITICAL_PRIORITY: '0 4px 12px rgba(232, 17, 35, 0.4)'
} as const;

// Carousel configuration
export const CAROUSEL_CONFIG = {
  AUTOPLAY_INTERVAL_MS: 5000,
  MIN_INTERVAL: 1000,
  MAX_INTERVAL: 60000
} as const;

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

export const UI_CONFIG = {
  FADE_IN_DURATION_MS: 300,
  FADE_OUT_DURATION_MS: 200,

  AUTO_SAVE_INTERVAL_MS: 30000, // 30 seconds

  TOAST_DURATION_MEDIUM_MS: 5000,
  TOAST_DURATION_LONG_MS: 8000,

  PREVIEW_DESCRIPTION_LENGTH: 150,
  CARD_TITLE_LENGTH: 60,
  LIST_ITEM_DESCRIPTION_LENGTH: 100
} as const;

export const ALERT_TYPE_DEFAULTS = {
  DEFAULT_TYPE: 'Info',
  DEFAULT_ICON: 'Info',
  DEFAULT_BACKGROUND_COLOR: '#0078d4',
  DEFAULT_TEXT_COLOR: '#ffffff'
} as const;

// Default Alert Type Name when none is specified
export const DEFAULT_ALERT_TYPE_NAME = "Info";

export const DEFAULT_ALERT_TYPES: IAlertType[] = [
  {
    name: "Info",
    iconName: "Info",
    backgroundColor: "#389899",
    textColor: "#ffffff",
    additionalStyles: "",
    defaultPriority: AlertPriority.Medium,
    priorityStyles: {
      [AlertPriority.Critical]: "border: 4px solid #E81123;",
      [AlertPriority.High]: "border: 3px solid #EA4300;",
      [AlertPriority.Medium]: "border: 2px solid #0078d4;",
      [AlertPriority.Low]: "border: 1px solid #107c10;"
    }
  },
  {
    name: "Warning",
    iconName: "Warning",
    backgroundColor: "#f1c40f",
    textColor: "#000000",
    additionalStyles: "",
    defaultPriority: AlertPriority.Medium,
    priorityStyles: {
      [AlertPriority.Critical]: "border: 4px solid #E81123;",
      [AlertPriority.High]: "border: 3px solid #EA4300;",
      [AlertPriority.Medium]: "border: 2px solid #f1c40f;",
      [AlertPriority.Low]: "border: 1px solid #107c10;"
    }
  },
  {
    name: "Maintenance",
    iconName: "ConstructionCone",
    backgroundColor: "#afd6d6",
    textColor: "#000000",
    additionalStyles: "",
    defaultPriority: AlertPriority.Low,
    priorityStyles: {
      [AlertPriority.Critical]: "border: 4px solid #E81123;",
      [AlertPriority.High]: "border: 3px solid #EA4300;",
      [AlertPriority.Medium]: "border: 2px solid #0078d4;",
      [AlertPriority.Low]: "border: 1px solid #afd6d6;"
    }
  },
  {
    name: "Interruption",
    iconName: "Error",
    backgroundColor: "#c54644",
    textColor: "#ffffff",
    additionalStyles: "",
    defaultPriority: AlertPriority.High,
    priorityStyles: {
      [AlertPriority.Critical]: "border: 4px solid #8B0000;",
      [AlertPriority.High]: "border: 3px solid #c54644;",
      [AlertPriority.Medium]: "border: 2px solid #EA4300;",
      [AlertPriority.Low]: "border: 1px solid #f1c40f;"
    }
  }
];

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

export const ALERT_STATUS = {
  ACTIVE: 'Active',
  EXPIRED: 'Expired',
  SCHEDULED: 'Scheduled',
  DRAFT: 'Draft'
} as const;

export const AUTO_SAVE_PREFIXES = {
  LOWERCASE: '[auto-saved]',
  CAPITALIZED: '[Auto-saved]'
} as const;

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

export const REGEX_PATTERNS = {
  EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
  URL: /^https?:\/\/.+/i,
  GUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
  LANGUAGE_FIELD: /^(Title|Description|LinkDescription)_[A-Z]{2}$/,
  HEX_COLOR: /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/
} as const;


export type ListName = typeof LIST_NAMES[keyof typeof LIST_NAMES];
export type FieldName = typeof FIELD_NAMES[keyof typeof FIELD_NAMES];
export type AlertStatus = typeof ALERT_STATUS[keyof typeof ALERT_STATUS];
export type LanguageCode = keyof typeof LANGUAGE_CODES;
export type LanguageSuffix = typeof LANGUAGE_CODES[LanguageCode];
