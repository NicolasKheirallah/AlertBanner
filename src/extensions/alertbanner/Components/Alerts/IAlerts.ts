import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

export interface IGraphListItem {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  createdBy: {
    user: {
      displayName: string;
      email?: string;
      userPrincipalName?: string;
      id?: string;
    };
  };
  // Retained 'any' here as removing it triggers a massive codebase-wide refactor
  // across all Services and Utils to cast dynamic SharePoint list responses.
  fields: { [key: string]: any };
}
// IAlertItem definition moved here to avoid circular dependencies
export interface IAlertItem {
  id: string;
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  targetUsers?: IPersonField[]; // People/Groups who can see this alert. If empty, everyone sees it
  notificationType: NotificationType;
  linkUrl?: string;
  linkDescription?: string;
  targetSites: string[];
  status: "Active" | "Expired" | "Scheduled" | "Draft";
  createdDate: string;
  createdBy: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  metadata?: unknown;
  // New language and classification properties
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageGroup?: string;
  availableForAll?: boolean;
  translationStatus?: TranslationStatus;
  languageContent?: ILanguageContent[];
  // Approval Workflow
  contentStatus?: ContentStatus;
  reviewer?: IPersonField[];
  reviewNotes?: string;
  submittedDate?: string;
  reviewedDate?: string;
  // Attachments support
  attachments?: {
    fileName: string;
    serverRelativeUrl: string;
    size?: number;
  }[];
  modified?: string;
  sortOrder?: number;
  // Store the original SharePoint list item for multi-language access
  _originalListItem?: IAlertListItem;
}

export interface IMultiLanguageContent {
  [languageCode: string]: string;
}

export interface ILanguageContent {
  language: TargetLanguage;
  title: string;
  description: string;
  linkDescription?: string;
  availableForAll?: boolean; // If true, this version can be shown to users of other languages
  translationStatus?: TranslationStatus;
}

export interface IMultiLanguageAlert {
  baseAlert: Omit<IAlertItem, "title" | "description" | "linkDescription">;
  content: ILanguageContent[];
  languageGroup: string;
}

export interface IAlertListItem {
  Id: number;
  Title: string;
  Description: string;
  AlertType: string;
  Priority: string;
  IsPinned: boolean;
  NotificationType: string;
  LinkUrl?: string;
  LinkDescription?: string;
  TargetSites: string;
  Status: string;
  Created: string;
  Author: {
    Title: string;
    Email?: string;
  };
  Modified?: string;
  ScheduledStart?: string;
  ScheduledEnd?: string;
  Metadata?: string;
  ItemType?: string;
  TargetLanguage?: string;
  LanguageGroup?: string;
  AvailableForAll?: boolean;
  TranslationStatus?: string;
  ContentStatus?: string;
  Reviewer?: IPersonField[];
  ReviewNotes?: string;
  SubmittedDate?: string;
  ReviewedDate?: string;

  // Targeting
  TargetUsers?: IPersonField[]; // SharePoint People/Groups field data
}

export enum AlertPriority {
  Low = "low",
  Medium = "medium",
  High = "high",
  Critical = "critical",
}

export enum NotificationType {
  None = "none",
  Browser = "browser",
  Email = "email",
  Both = "both",
}

export enum ContentType {
  Alert = "alert",
  Template = "template",
  Draft = "draft",
}

export enum TranslationStatus {
  Draft = "Draft",
  InReview = "InReview",
  Approved = "Approved",
}

export enum ContentStatus {
  Draft = "Draft",
  PendingReview = "PendingReview",
  Approved = "Approved",
  Rejected = "Rejected",
}

export enum TargetLanguage {
  EnglishUS = "en-us",
  FrenchFR = "fr-fr",
  GermanDE = "de-de",
  SpanishES = "es-es",
  SwedishSE = "sv-se",
  FinnishFI = "fi-fi",
  DanishDK = "da-dk",
  NorwegianNO = "nb-no",
  All = "all", // For items that should show to all languages
}

// Interface for SharePoint Person field data
export interface IPersonField {
  id: string; // User/Group ID
  displayName: string; // Display name
  email?: string; // Email address (for users)
  loginName?: string; // Login name
  isGroup: boolean; // Whether this is a group or individual user
}

// Interface for targeting rules
export interface ITargetingRule {
  // Support for People fields
  targetUsers?: IPersonField[]; // Individual users from People field
  targetGroups?: IPersonField[]; // SharePoint groups from People field

  // Legacy targeting with string arrays
  audiences?: string[];

  // Operation to apply
  operation: "anyOf" | "allOf" | "noneOf";
}

export interface IAlertsBannerApplicationCustomizerProperties {
  alertTypesJson: string; // Property to hold the alert types JSON
  userTargetingEnabled: boolean; // Enable user targeting feature
  notificationsEnabled: boolean; // Enable notifications feature
  enableTargetSite: boolean; // Enable targeting specific sites (default: false)
  emailServiceAccount?: string; // Service account email for Graph API sendMail
  copilotEnabled?: boolean; // Enable Copilot features
}

export interface IAlertsProps {
  siteIds?: string[];
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  alertTypesJson: string; // Property to receive the alert types JSON
  userTargetingEnabled?: boolean;
  notificationsEnabled?: boolean;
  enableTargetSite?: boolean; // Control visibility of target site selector
  emailServiceAccount?: string; // Service account email for email notifications
  copilotEnabled?: boolean; // Enable Copilot features
  onSettingsChange?: (settings: {
    alertTypesJson: string;
    userTargetingEnabled: boolean;
    notificationsEnabled: boolean;
    enableTargetSite: boolean;
    emailServiceAccount?: string;
    copilotEnabled?: boolean;
  }) => void;
}

export interface IAlertsState {
  alerts: IAlertItem[];
  alertTypes: { [key: string]: IAlertType };
  isLoading: boolean;
  hasError: boolean;
  errorMessage?: string;
  userDismissedAlerts: string[];
  userHiddenAlerts: string[];
  currentIndex: number;
  isInEditMode: boolean;
}

export interface IUser {
  id: string;
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
  userGroups?: string[];
}

// IAlertRichMedia removed - using description field for all content

export interface IQuickAction {
  label: string;
  actionType: "link" | "dismiss" | "acknowledge" | "custom";
  url?: string; // For link type actions
  callback?: string; // Function name to execute for custom actions
  icon?: string; // Icon name for the action button
}

export interface IPriorityColorConfig {
  borderColor?: string; // Color for the left border
}

export interface IAlertType {
  id?: string; // SharePoint item ID (for lookup field reference)
  name: string;
  iconName: string;
  backgroundColor: string;
  textColor: string;
  additionalStyles?: string;
  defaultPriority?: AlertPriority; // Optional default priority to be selected when this type is chosen
  priorityStyles?: { [key in AlertPriority]?: string }; // Different styles based on priority (CSS)
  priorityColors?: { [key in AlertPriority]?: IPriorityColorConfig }; // Structured color config per priority
}
