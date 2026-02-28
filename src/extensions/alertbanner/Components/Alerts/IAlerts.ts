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
  fields: { [key: string]: any };
}
export interface IAlertItem {
  id: string;
  title: string;
  description: string;
  AlertType: string;
  priority: AlertPriority;
  isPinned: boolean;
  targetUsers?: IPersonField[];
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
  contentType: ContentType;
  targetLanguage: TargetLanguage;
  languageGroup?: string;
  availableForAll?: boolean;
  translationStatus?: TranslationStatus;
  languageContent?: ILanguageContent[];
  contentStatus?: ContentStatus;
  reviewer?: IPersonField[];
  reviewNotes?: string;
  submittedDate?: string;
  reviewedDate?: string;
  attachments?: {
    fileName: string;
    serverRelativeUrl: string;
    size?: number;
  }[];
  modified?: string;
  sortOrder?: number;
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
  availableForAll?: boolean;
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

  TargetUsers?: IPersonField[];
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
  All = "all",
}

export interface IPersonField {
  id: string;
  displayName: string;
  email?: string;
  loginName?: string;
  isGroup: boolean;
}

export interface ITargetingRule {
  targetUsers?: IPersonField[];
  targetGroups?: IPersonField[];

  audiences?: string[];

  operation: "anyOf" | "allOf" | "noneOf";
}

export interface IAlertsBannerApplicationCustomizerProperties {
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  enableTargetSite: boolean;
  emailServiceAccount?: string;
  copilotEnabled?: boolean;
}

export interface IAlertsProps {
  siteIds?: string[];
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  alertTypesJson: string;
  userTargetingEnabled?: boolean;
  notificationsEnabled?: boolean;
  enableTargetSite?: boolean;
  emailServiceAccount?: string;
  copilotEnabled?: boolean;
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

export interface IQuickAction {
  label: string;
  actionType: "link" | "dismiss" | "acknowledge" | "custom";
  url?: string;
  callback?: string;
  icon?: string;
}

export interface IPriorityColorConfig {
  borderColor?: string;
}

export interface IAlertType {
  id?: string;
  name: string;
  iconName: string;
  backgroundColor: string;
  textColor: string;
  additionalStyles?: string;
  defaultPriority?: AlertPriority;
  priorityStyles?: { [key in AlertPriority]?: string };
  priorityColors?: { [key in AlertPriority]?: IPriorityColorConfig };
}
