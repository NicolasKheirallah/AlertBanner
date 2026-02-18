define([], function () {
  return {
    Title: "AlertBannerApplicationCustomizer",

    // General UI
    Save: "Save",
    Cancel: "Cancel",
    Close: "Close",
    Edit: "Edit",
    Delete: "Delete",
    Add: "Add",
    Update: "Update",
    Create: "Create",
    Remove: "Remove",
    Yes: "Yes",
    No: "No",
    OK: "OK",
    Loading: "Loading...",
    Error: "Error",
    Success: "Success",
    Warning: "Warning",
    Info: "Info",

    // Language
    Language: "Language",
    SelectLanguage: "Select Language",
    ChangeLanguage: "Change Language",

    // Time-related
    JustNow: "Just now",
    MinutesAgo: "{0} minutes ago",
    HoursAgo: "{0} hours ago",
    DaysAgo: "{0} days ago",

    // Alert Settings
    AlertSettings: "Alert Settings",
    AlertSettingsTitle: "Alert Banner Settings",
    AlertSettingsDescription:
      "Configure the alert banner settings. These changes will be applied site-wide.",
    ConfigureAlertBannerSettings: "Configure Alert Banner Settings",
    Features: "Features",
    EnableUserTargeting: "Enable User Targeting",
    EnableUserTargetingDescription:
      "Allow alerts to target specific users or groups",
    EnableNotifications: "Enable Notifications",
    EnableNotificationsDescription:
      "Send browser notifications for critical alerts",
    EnableRichMedia: "Enable Rich Media",
    EnableRichMediaDescription:
      "Support images, videos, and rich content in alerts",
    AlertTypesConfiguration: "Alert Types Configuration",
    AlertTypesConfigurationDescription:
      "Configure the available alert types (JSON format):",
    AlertTypesPlaceholder: "Enter alert types JSON configuration...",
    AlertTypesHelpText:
      "Each alert type should have: name, iconName, backgroundColor, textColor, additionalStyles, priorityStyles, and optionally defaultPriority (low, medium, high, critical)",
    SaveSettings: "Save Settings",
    InvalidJSONError:
      "Invalid JSON format in Alert Types configuration. Please check your syntax.",
    ManageAlerts: "Manage Alerts",
    AlertTypesTabTitle: "Alert Types",
    SettingsTabTitle: "Settings",
    SettingsSaveAll: "Save Changes",
    SettingsSaving: "Saving...",
    SettingsDiscardChanges: "Discard Changes",
    SettingsKeepEditing: "Keep Editing",
    SettingsUnsavedChangesTitle: "Discard Unsaved Settings?",
    SettingsUnsavedChangesMessage:
      "You have unsaved settings changes. Leave this tab and discard them?",
    SettingsUnsavedChangesLabel: "Unsaved changes",
    SettingsLastSavedLabel: "Last saved at {0}",
    SettingsPreflightErrorsTitle: "Fix these issues before saving",
    SettingsPreflightWarningsTitle: "Review before saving",
    SettingsValidationCarouselNumber:
      "Carousel interval must be a whole number.",
    SettingsValidationCarouselRange:
      "Carousel interval must be between {0} and {1} seconds.",
    SettingsValidationEmail:
      "Email Service Account must be a valid email address.",
    SettingsValidationAlertTypesJson:
      "Alert type configuration JSON is invalid.",
    SettingsWarningNotificationsWithoutEmail:
      "Browser notifications are enabled without an email service account.",
    SettingsReadOnlyWarning:
      "You can review settings, but saving requires Site Admin permissions.",
    SettingsSectionPermissionsTitle: "Permissions & Readiness",
    SettingsSectionAudienceTitle: "Audience & Targeting",
    SettingsSectionIntegrationsTitle: "Integrations",
    SettingsSectionExperienceTitle: "Experience",
    SettingsSectionLocalizationTitle: "Localization",
    SettingsSectionSetupTitle: "SharePoint Setup",
    SettingsSectionMaintenanceTitle: "Maintenance",
    SettingsEnableTargetSitesLabel: "Enable Target Site Selection",
    SettingsEnableTargetSitesDescription:
      "Allow administrators to target specific sites for each alert.",
    SettingsEnableNotificationsDescription:
      "Send browser notifications for high-priority alerts. Keep disabled if you want banners only.",
    SettingsEmailServiceAccountLabel: "Email Service Account",
    SettingsEmailServiceAccountPlaceholder:
      "alerts-noreply@yourtenant.onmicrosoft.com",
    SettingsEmailServiceAccountDescription:
      "Service account used for Graph-based email notifications (Mail.Send required).",
    SettingsSendTestEmailButton: "Send Test Email",
    SettingsGraphUnavailableError:
      "Graph client is unavailable in this context.",
    SettingsTestEmailSuccess: "Test email sent. Check your inbox.",
    SettingsTestEmailFailed:
      "Failed to send test email. Verify Mail.Send permission and account configuration.",
    SettingsCarouselEnabledLabel: "Enable Carousel Auto-Rotation",
    SettingsCarouselEnabledDescription:
      "Automatically rotate between alerts when multiple alerts are visible.",
    SettingsCarouselIntervalLabel: "Carousel Timer (seconds)",
    SettingsCarouselIntervalDescription:
      "Time in seconds between automatic transitions.",
    SettingsCheckingLists: "Checking SharePoint lists...",
    SettingsMissingListsDescription:
      "Some required lists are missing and need to be created on this site.",
    SettingsCurrentSiteLabel: "Current Site:",
    SettingsMissingAlertsListItem:
      "Alerts list is missing (stores alert content).",
    SettingsMissingTypesListItem:
      "AlertBannerTypes list is missing (stores styling definitions).",
    SettingsInitialLanguagesTitle: "Select Languages for Initial Setup",
    SettingsInitialLanguagesDescription:
      "Choose languages to provision now. You can add more later.",
    SettingsInitialLanguagesHint:
      "English is always included and cannot be removed.",
    SettingsCreatingLists: "Creating Lists...",
    SettingsCreateMissingLists: "Create Missing Lists",
    SettingsCreateMissingListsHelp:
      "Creates only missing lists on the current site.",
    SettingsSetupCompleteTitle: "Setup Complete",
    SettingsSetupCompleteDescription:
      "All required SharePoint lists are configured and ready to use.",
    SettingsListMaintenanceTitle: "List Maintenance",
    SettingsListMaintenanceDescription:
      "Repair list schema by removing obsolete columns and adding missing ones.",
    SettingsStorageDescription:
      "Manage local cache and advanced settings import/export actions.",
    SettingsClearCache: "Clear Alert Cache",
    SettingsCacheCleared: "Alert cache cleared.",
    SettingsResetCarousel: "Reset Carousel Settings",
    SettingsCarouselResetSuccess: "Carousel settings reset to defaults.",
    SettingsShowAdvanced: "Show Advanced",
    SettingsHideAdvanced: "Hide Advanced",
    SettingsExportButton: "Export Settings JSON",
    SettingsImportButton: "Import Settings JSON",
    SettingsResetToDefaults: "Reset to Defaults",
    SettingsExportSuccess: "Settings exported.",
    SettingsExportFailed: "Failed to export settings.",
    SettingsImportSuccess: "Settings imported into the draft editor.",
    SettingsImportFailed: "Failed to import settings JSON.",
    SettingsProgressCheckListsName: "Checking existing lists",
    SettingsProgressCheckListsDescription:
      "Verifying what lists need to be created.",
    SettingsProgressCreateListsName: "Creating SharePoint lists",
    SettingsProgressCreateListsDescription:
      "Setting up Alerts and AlertBannerTypes lists.",
    SettingsProgressAddLanguageName: "Adding {0} language support",
    SettingsProgressAddLanguageDescription:
      "Creating language-specific columns for {0}.",
    SettingsProgressFinalizeName: "Finalizing setup",
    SettingsProgressFinalizeDescription:
      "Completing configuration and verification.",
    SettingsListsAlreadyExistMessage:
      "All required lists already exist on this site.",
    SettingsListsAlreadyExistTitle: "Lists Already Exist",
    SettingsListsCreatedTitle: "Lists Created Successfully",
    SettingsAdditionalInformationTitle: "Additional Information",
    SettingsContactAdministrator: "Contact Administrator",

    // Create Alert Form
    CreateAlertSectionContentClassificationTitle: "Content Classification",
    ContentTypeLabel: "Content Type",
    CreateAlertSectionContentClassificationDescription:
      "Choose how this alert content should be classified.",
    CreateAlertLanguageConfigurationLabel: "Language Configuration",
    CreateAlertSingleLanguageButton: "Single Language",
    CreateAlertMultiLanguageButton: "Multi-language",
    CreateAlertSectionUserTargetingTitle: "User Targeting",
    CreateAlertPeoplePickerLabel: "Target specific users or groups",
    CreateAlertPeoplePickerDescription:
      "Leave empty to show to everyone. Select users or groups to restrict visibility.",
    CreateAlertSectionLanguageTargetingTitle: "Language Targeting",
    CreateAlertTargetLanguageLabel: "Target Language",
    CreateAlertSectionLanguageTargetingDescription:
      "Select the language that should receive this alert.",
    CreateAlertSectionBasicInformationTitle: "Basic Information",
    CreateAlertTitlePlaceholder: "Enter alert title",
    CreateAlertTitleDescription:
      "Provide a short, descriptive title for your alert.",
    CreateAlertDescriptionPlaceholder: "Enter alert details",
    CreateAlertDescriptionHelp:
      "Use rich text to provide the full alert message. You can include links and formatting.",
    CreateAlertConfigurationSectionTitle: "Alert Configuration",
    CreateAlertConfigurationDescription:
      "Choose the visual style and importance level",
    CreateAlertPriorityLabel: "Priority Level",
    CreateAlertPriorityDescription:
      "This affects the visual styling and user attention level",
    CreateAlertPinLabel: "Pin Alert",
    CreateAlertPinDescription:
      "Pinned alerts stay at the top and are harder to dismiss",
    CreateAlertNotificationLabel: "Notification Type",
    CreateAlertNotificationDescription:
      "How users will be notified about this alert",
    CreateAlertActionLinkSectionTitle: "Action Link (Optional)",
    CreateAlertLinkUrlLabel: "Link URL",
    CreateAlertLinkUrlPlaceholder: "https://example.com/more-info",
    CreateAlertLinkUrlDescription:
      "Optional link for users to get more information or take action",
    CreateAlertLinkDescriptionLabel: "Link Description",
    CreateAlertLinkDescriptionPlaceholder: "Learn More",
    CreateAlertLinkDescriptionDescription:
      "Text that will appear on the action button",
    CreateAlertLinkDescriptionInfo:
      "Link descriptions will be configured per language in the Multi-Language Content section above.",
    CreateAlertTargetSitesSectionTitle: "Target Sites",
    CreateAlertSchedulingSectionTitle: "Scheduling (Optional)",
    CreateAlertStartDateLabel: "Start Date & Time",
    CreateAlertStartDateDescription:
      "When should this alert become visible? Leave empty to show immediately.",
    CreateAlertEndDateLabel: "End Date & Time",
    CreateAlertEndDateDescription:
      "When should this alert automatically hide? Leave empty to keep it visible until manually removed.",
    CreateAlertPrimaryButtonLoading: "Creating Alert...",
    CreateAlertHidePreview: "Hide Preview",
    CreateAlertShowPreview: "Show Preview",
    CreateAlertAutoSaving: "💾 Auto-saving...",
    CreateAlertAutoSavePending: "Changes pending auto-save...",
    CreateAlertAutoSaveRetrying: "Auto-save failed. Retrying automatically.",
    CreateAlertAutoSavedAt: "✓ Auto-saved at {0}",
    CreateAlertSaveDraftButtonLabel: "Save as Draft",
    CreateAlertResetFormButtonLabel: "Reset Form",
    CreateAlertTemplatesButtonLabel: "Templates",
    CreateAlertPreviousAlertsButtonLabel: "Previous Alerts ({0})",
    CreateAlertDraftsButtonLabel: "My Drafts ({0})",
    CreateAlertNoDraftsTitle: "No drafts yet",
    CreateAlertNoPreviousAlertsTitle: "No recent alerts",
    CreateAlertDraftsEmptyMessage:
      "No drafts found. Save your work as a draft to continue later!",
    CreateAlertPreviousAlertsEmptyMessage:
      "No recent alerts found on this site yet.",
    CreateAlertLoadDraftButton: "Load Draft",
    CreateAlertUsePreviousButton: "Use Alert",
    CreateAlertStartFromScratchButton: "Start from Scratch",
    CreateAlertDraftTitleError:
      "Draft must have a title (at least 3 characters)",
    CreateAlertAlertTypeError:
      "Please select an alert type before saving draft",
    CreateAlertMultiLanguageSuccessStatus:
      "Multi-language alert created ({0} variants)",
    CreateAlertMultiLanguageSuccessTitle: "Alerts Created",
    CreateAlertCreationSuccessStatus: "Alert Created",
    CreateAlertCreationSuccessTitle: "Alert Created",
    CreateAlertCreationErrorStatus: "Creation Error",
    CreateAlertCreationErrorTitle: "Creation Failed",
    CreateAlertCreationResultsHeading: "Creation Results",
    CreateAlertCreationResultSuccess: "✅ Created successfully",
    CreateAlertCreationResultError: "❌ {0}",
    CreateAlertCreationRetryButton: "Retry Create Alert",
    CreateAlertDraftSaveSuccessStatus: "Draft Saved Successfully",
    CreateAlertDraftSaveErrorStatus: "Draft Save Error",
    CreateAlertUnknownError: "Unknown error occurred",
    CreateAlertLanguageRequired: "At least one language must be configured",
    CreateAlertLanguageTitleRequired: "Title is required for {0}",
    CreateAlertLanguageDescriptionRequired:
      "Description is required for {0}",
    CreateAlertLanguageLinkDescriptionRequired:
      "Link description is required for {0} when URL is provided",
    CreateAlertLanguageAtLeastOneComplete:
      "At least one language must have complete content (title and description)",
    CreateAlertDefaultLanguageRequired:
      "Default language ({0}) must have complete content",
    DuplicateLanguagesNotAllowed: "Duplicate language variants detected: {0}",
    CreateAlertPreviewLanguageLabel: "Preview Language",
    CreateAlertMultiLanguagePreviewTitle: "Multi-language Alert Title",
    CreateAlertMultiLanguagePreviewDescription:
      "Multi-language alert description will appear here...",
    CreateAlertMultiLanguagePreviewHeading: "Multi-language Alert",
    CreateAlertMultiLanguagePreviewCurrentLanguage: "Currently previewing: {0}",
    CreateAlertMultiLanguagePreviewAvailableLanguages:
      "Available in {0} language(s): {1}",
    CreateAlertPreviewDescriptionFallback:
      "Alert description will appear here...",
    CreateAlertLinkPreviewFallback: "Learn More",
    CreateAlertPriorityLowDescription: "General updates or low urgency items.",
    CreateAlertPriorityMediumDescription:
      "Important information that should be seen soon.",
    CreateAlertPriorityHighDescription:
      "Time-sensitive alert that requires quick attention.",
    CreateAlertPriorityCriticalDescription:
      "Critical incident or outage that needs immediate action.",
    CreateAlertNotificationNoneDescription:
      "Do not send a notification—users will see the banner only.",
    CreateAlertNotificationBrowserDescription:
      "Send a browser notification alongside the banner.",
    CreateAlertNotificationEmailDescription:
      "Send an email to targeted users when the alert publishes.",
    CreateAlertNotificationBothDescription:
      "Send both browser and email notifications.",
    CreateAlertContentTypeAlertDescription: "📢 Alert - Live content for users",
    CreateAlertContentTypeTemplateDescription:
      "📄 Template - Reusable template for future alerts",
    CreateAlertContentTypeDraftDescription: "✏️ Draft - Work in progress",
    CreateAlertTargetLanguageAll: "All supported languages",
    CreateAlertWizardContentStep: "Content",
    CreateAlertWizardAudienceStep: "Actions & Media",
    CreateAlertWizardPublishStep: "Publish",
    CreateAlertWizardStepComplete: "Ready",
    CreateAlertWizardStepIncomplete: "Needs input",
    CreateAlertWizardBackButton: "Back",
    CreateAlertWizardNextButton: "Next",
    CreateAlertPublishSummaryTitle: "Confirm Alert Creation",
    CreateAlertSchedulePresetNow: "Start Now",
    CreateAlertSchedulePresetEndOfDay: "End of Day",
    CreateAlertSchedulePresetPlusWeek: "End in 1 Week",
    CreateAlertPriorityInheritedMessage:
      "Priority default was applied from alert type: {0}.",
    CreateAlertImageManagerHint:
      "Set a title or enable multi-language to initialize an image folder.",
    CreateAlertCopilotUnavailableMessage:
      "Copilot is enabled but unavailable for this account or tenant.",
    CreateAlertUnsavedChangesTitle: "Discard Unsaved Changes?",
    CreateAlertUnsavedChangesMessage:
      "You have unsaved create-alert changes. Leave this tab and discard them?",
    CreateAlertDiscardChangesButton: "Discard Changes",
    CreateAlertKeepEditingButton: "Keep Editing",

    // Alert Management
    CreateAlert: "Create Alert",
    EditAlert: "Edit Alert",
    DeleteAlert: "Delete Alert",
    AlertTitle: "Alert Title",
    AlertDescription: "Description",
    AlertType: "Alert Type",
    Priority: "Priority",
    Status: "Status",
    TargetSites: "Target Sites",
    LinkUrl: "Link URL",
    LinkDescription: "Link Description",
    ScheduledStart: "Scheduled Start",
    ScheduledEnd: "Scheduled End",
    IsPinned: "Pinned",
    NotificationType: "Notification Type",
    ManageAlertsSubtitle:
      "View, edit, and manage existing alerts across your sites",
    ManageAlertsUnsavedChangesTitle: "Discard Unsaved Changes?",
    ManageAlertsUnsavedChangesMessage:
      "You have unsaved manage-alert changes. Leave this view and discard them?",
    ManageAlertsDiscardChangesButton: "Discard Changes",
    ManageAlertsKeepEditingButton: "Keep Editing",
    ManageAlertsEditSelectedButtonLabel: "Edit Selected",
    ManageAlertsDeleteSelectedButtonLabel: "Delete Selected ({0})",
    ManageAlertsRefreshingLabel: "Refreshing...",
    ManageAlertsHideFiltersLabel: "Hide",
    ManageAlertsShowFiltersLabel: "Show",
    ManageAlertsFiltersLabel: "Filters",
    ManageAlertsLoadingTitle: "Loading alerts...",
    ManageAlertsLoadingSubtitle: "Please wait while we fetch your alerts",
    ManageAlertsEmptyTitle: "No Alerts Found",
    ManageAlertsEmptyDescription:
      "No alerts are currently available. This might be because:",
    ManageAlertsEmptyReasonListsMissing:
      "The Alert Banner lists haven't been created yet",
    ManageAlertsEmptyReasonNoAccess: "You don't have access to the lists",
    ManageAlertsEmptyReasonNoAlerts: "No alerts have been created yet",
    ManageAlertsEmptyCreateFirstButton: "Create First Alert",
    ManageAlertsFilterSummaryCount: "Showing {0} of {1} items",
    ManageAlertsFilterSummaryLanguageGroups: "{0} multi-language groups",
    ManageAlertsFilterAlertsLabel: "📢 Alerts ({0})",
    ManageAlertsNoDraftsTitle: "No drafts yet",
    ManageAlertsNoTemplatesTitle: "No templates available",
    ManageAlertsFilterDraftsLabel: "✏️ Drafts ({0})",
    ManageAlertsFilterTemplatesLabel: "📄 Templates ({0})",
    ManageAlertsFilterAllLabel: "📋 All ({0})",
    ManageAlertsSearchLabel: "Search alerts",
    ManageAlertsSearchPlaceholder:
      "🔍 Search alerts, drafts, and templates by title, description, type, priority, or author...",
    ManageAlertsContentTypeAllOption: "📋 All Types",
    ManageAlertsContentTypeAlertOption: "📢 Alerts",
    ManageAlertsContentTypeDraftOption: "✏️ Drafts",
    ManageAlertsContentTypeTemplateOption: "📄 Templates",
    ManageAlertsPriorityAllOption: "🔘 All Priorities",
    ManageAlertsAlertTypeAllOption: "🎨 All Types",
    ManageAlertsStatusAllOption: "⚪ All Statuses",
    ManageAlertsLanguageAllOption: "🌐 All Languages",
    ManageAlertsLanguageMultiOption: "🌍 Multi-Language",
    ManageAlertsNotificationAllOption: "📧 All Notification Types",
    ManageAlertsNotificationNoneOption: "🚫 None - Banner only",
    ManageAlertsNotificationBrowserOption: "🌐 Browser - Banner display",
    ManageAlertsNotificationEmailOption: "📧 Email only - No banner",
    ManageAlertsNotificationBothOption: "📧🌐 Browser + Email",
    ManageAlertsCreatedDateFilterLabel: "Created Date",
    ManageAlertsDateAllOption: "📅 All Dates",
    ManageAlertsDateTodayOption: "📅 Today",
    ManageAlertsDateWeekOption: "📅 This Week",
    ManageAlertsDateMonthOption: "📅 This Month",
    ManageAlertsDateCustomOption: "📅 Custom Range",
    ManageAlertsFromDateLabel: "From Date",
    ManageAlertsToDateLabel: "To Date",
    ManageAlertsClearFiltersButton: "Clear All Filters",
    ManageAlertsSelectItemLabel: "Select alert {0}",
    ManageAlertsNoTitleFallback: "[No Title Available]",
    ManageAlertsNoDescriptionFallback: "[No Description Available]",
    ManageAlertsTemplateBadge: "📄 Template",
    ManageAlertsAlertBadge: "📢 Alert",
    ManageAlertsLanguageBadge: "{0} languages",
    ManageAlertsMultiLanguageList: "🌐 Multi-language ({0})",
    ManageAlertsAllLanguagesLabel: "🌐 All Languages",
    ManageAlertsPublishButtonLabel: "Publish",
    ManageAlertsCreatedLabel: "Created",
    ManageAlertsContentClassificationDescription:
      "Choose whether this is a live alert, draft, or reusable template",
    ManageAlertsLanguageConfigurationLabel: "Language Configuration",
    ManageAlertsSingleLanguageButton: "🌐 Single Language",
    ManageAlertsMultiLanguageButton: "🗣️ Multi-Language",
    ManageAlertsTargetLanguageDescription:
      "Choose which language audience this alert targets",
    ManageAlertsDescriptionPlaceholder:
      "Provide detailed information about the alert...",
    ManageAlertsDescriptionHelp:
      "Use the toolbar to format your message with rich text, links, lists, emojis, and images. Manage attachments below.",
    ManageAlertsUploadedImagesSectionTitle: "Uploaded Images",
    ManageAlertsAlertTypeDescription:
      "Select the visual style and category for this alert",
    ManageAlertsPriorityDescription:
      "Set the importance level—affects styling and user attention",
    ManageAlertsPinDescription:
      "Pinned alerts appear at the top and are more prominent",
    ManageAlertsNotificationMethodLabel: "Notification Method",
    ManageAlertsNotificationDescription:
      "Choose how users will be notified about this alert",
    ManageAlertsLinkDescriptionInfo:
      "💡 Link descriptions are configured per language in the Multi-Language Content section above.",
    ManageAlertsTargetSitesDescription:
      "Choose which SharePoint sites will display this alert.",
    ManageAlertsSchedulingDescription:
      "Configure when this alert will be visible. Leave dates empty for immediate activation and manual control.",
    ManageAlertsStartDateDescription: "When this alert becomes active",
    ManageAlertsEndDateDescription: "When this alert expires",
    ManageAlertsScheduleSummaryTitle: "Schedule Summary",
    ManageAlertsScheduleImmediate:
      "⚡ Alert is active immediately and requires manual deactivation.",
    ManageAlertsScheduleStartOnly:
      "📅 Alert activates on {0} and remains active until manually deactivated.",
    ManageAlertsScheduleEndOnly: "⏰ Alert is active immediately until {0}.",
    ManageAlertsScheduleWindow: "📅 Active from {0} to {1}.",
    ManageAlertsScheduleTimezone:
      "🌍 All times are in your local time zone ({0}).",
    ManageAlertsSavingChangesLabel: "Saving Changes...",
    ManageAlertsSaveChangesButton: "Save Changes",
    ManageAlertsLivePreviewTitle: "Live Preview",
    ManageAlertsMultiLanguagePreviewHeading: "Multi-language Preview",
    ManageAlertsMultiLanguagePreviewCurrentLanguage:
      "Currently previewing: {0}",
    ManageAlertsMultiLanguagePreviewAvailableLanguages:
      "Available in {0} language(s): {1}",
    ManageAlertsBulkDeleteDialogTitle: "Delete Alerts",
    ManageAlertsBulkDeleteDialogMessage:
      "Are you sure you want to delete {0} alert(s)? This action cannot be undone.",
    ManageAlertsBulkDeleteSuccessMessage: "Successfully deleted {0} alert(s)",
    ManageAlertsBulkDeleteSuccessTitle: "Alerts Deleted",
    ManageAlertsBulkDeleteErrorMessage:
      "Failed to delete some alerts. Please try again.",
    ManageAlertsBulkDeleteErrorTitle: "Deletion Failed",
    ManageAlertsDeleteDialogTitle: "Delete Alert",
    ManageAlertsDeleteDialogMessage:
      "Are you sure you want to delete \"{0}\"? This action cannot be undone.",
    ManageAlertsDeleteSuccessMessage: "Successfully deleted \"{0}\"",
    ManageAlertsDeleteSuccessTitle: "Alert Deleted",
    ManageAlertsDeleteErrorMessage: "Failed to delete alert. Please try again.",
    ManageAlertsDeleteErrorTitle: "Deletion Failed",
    ManageAlertsEditOpenErrorMessage:
      "Failed to open edit dialog. Please try again.",
    ManageAlertsEditOpenErrorTitle: "Edit Failed",
    ManageAlertsPublishDialogTitle: "Publish Draft",
    ManageAlertsPublishDialogMessage:
      "Publish \"{0}\" as a live alert? It will be visible to users immediately.",
    ManageAlertsPublishDialogConfirm: "Publish",
    ManageAlertsPublishSuccessMessage: "Successfully published \"{0}\"",
    ManageAlertsPublishSuccessTitle: "Draft Published",
    ManageAlertsPublishErrorMessage: "Failed to publish draft. Please try again.",
    ManageAlertsPublishErrorTitle: "Publish Failed",
    ManageAlertsSubmitDialogTitle: "Submit for Review",
    ManageAlertsSubmitDialogMessage: "Submit \"{0}\" for review?",
    ManageAlertsSubmitDialogConfirm: "Submit",
    ManageAlertsSubmitSuccessMessage:
      "Successfully submitted \"{0}\" for review",
    ManageAlertsSubmitSuccessTitle: "Alert Submitted",
    ManageAlertsSubmitErrorMessage: "Failed to submit alert. Please try again.",
    ManageAlertsSubmitErrorTitle: "Submission Failed",
    ManageAlertsApproveDialogTitle: "Approve Alert",
    ManageAlertsApproveDialogMessage:
      "Approve \"{0}\"? It will become visible to users based on scheduling.",
    ManageAlertsApproveDialogConfirm: "Approve",
    ManageAlertsApproveSuccessMessage: "Successfully approved \"{0}\"",
    ManageAlertsApproveSuccessTitle: "Alert Approved",
    ManageAlertsApproveErrorMessage: "Failed to approve alert. Please try again.",
    ManageAlertsApproveErrorTitle: "Approval Failed",
    ManageAlertsRejectDialogTitle: "Reject Alert",
    ManageAlertsRejectDialogMessage:
      "Enter rejection notes for \"{0}\" (optional):",
    ManageAlertsRejectDialogConfirm: "Reject",
    ManageAlertsRejectDialogLabel: "Notes",
    ManageAlertsRejectDialogPlaceholder: "Enter rejection notes...",
    ManageAlertsRejectSuccessMessage: "Rejected \"{0}\"",
    ManageAlertsRejectSuccessTitle: "Alert Rejected",
    ManageAlertsRejectErrorMessage: "Failed to reject alert. Please try again.",
    ManageAlertsRejectErrorTitle: "Rejection Failed",
    ManageAlertsUpdateSuccessMessage: "Alert updated successfully!",
    ManageAlertsUpdateSuccessTitle: "Update Complete",
    ManageAlertsUpdateErrorMessage: "Failed to update alert. Please try again.",
    ManageAlertsUpdateErrorTitle: "Update Failed",
    ManageAlertsApprovalStatusLabel: "Approval Status",
    ManageAlertsApprovalAnyStatusOption: "Any Status",
    ManageAlertsApprovalDraftOption: "✏️ Draft",
    ManageAlertsApprovalPendingOption: "⏳ Pending Review",
    ManageAlertsApprovalApprovedOption: "✅ Approved",
    ManageAlertsApprovalRejectedOption: "❌ Rejected",
    ManageAlertsSubmitForReviewTitle: "Submit for Review",
    ManageAlertsSubmitForReviewButton: "Submit",
    ManageAlertsApproveButton: "Approve",
    ManageAlertsRejectButton: "Reject",
    ManageAlertsPinnedBadge: "📌 PINNED",
    ManageAlertsContentStatusPending: "⏳ Pending",
    ManageAlertsContentStatusRejected: "❌ Rejected",
    ManageAlertsContentStatusDraft: "✏️ Draft",

    // Priority Levels
    PriorityLow: "Low",
    PriorityMedium: "Medium",
    PriorityHigh: "High",
    PriorityCritical: "Critical",

    // Status Types
    StatusActive: "Active",
    StatusExpired: "Expired",
    StatusScheduled: "Scheduled",
    StatusInactive: "Inactive",

    // Notification Types
    NotificationNone: "None",
    NotificationBrowser: "Browser",
    NotificationEmail: "Email",
    NotificationBoth: "Both",

    // Alert Types
    AlertTypeInfo: "Info",
    AlertTypeWarning: "Warning",
    AlertTypeIconLabel: "Icon",
    AlertTypeChooseIconButton: "Choose Fluent UI Icon",
    AlertTypeSelectedIconNameLabel: "Selected Icon Name",
    AlertTypeSelectedIconNamePlaceholder: "Info",
    AlertTypeSelectedIconNameDescription:
      "Selected Fluent UI icon name. You can still type a custom icon name.",
    AlertTypeIconDialogTitle: "Select a Fluent UI Icon",
    AlertTypeIconSearchLabel: "Search icons",
    AlertTypeIconSearchPlaceholder: "Type icon name (e.g., Shield, Calendar, Megaphone)",
    AlertTypeIconResultsCount: "{0} icon(s)",
    AlertTypeIconNoResultsTitle: "No matching icons",
    AlertTypeIconNoResultsDescription:
      "Try a different search term to find a Fluent UI icon.",
    AlertTypeMaintenance: "Maintenance",
    AlertTypeInterruption: "Interruption",

    // User Interface
    ShowMore: "Show More",
    ShowLess: "Show Less",
    ViewDetails: "View Details",
    Expand: "Expand",
    Collapse: "Collapse",
    Preview: "Preview",
    Templates: "Templates",
    CustomizeColors: "Customize Colors",

    // Site Selection
    SelectSites: "Select Sites",
    CurrentSite: "Current Site",
    AllSites: "All Sites",
    HubSites: "Hub Sites",
    RecentSites: "Recent Sites",
    FollowedSites: "Followed Sites",

    // Permissions and Errors
    InsufficientPermissions: "Insufficient permissions to perform this action",
    PermissionDeniedCreateLists:
      "User lacks permissions to create SharePoint lists",
    PermissionDeniedAccessLists:
      "User lacks permissions to access SharePoint lists",
    ListsNotFound: "SharePoint lists do not exist and cannot be created",
    InitializationFailed: "Failed to initialize SharePoint connection",
    ConnectionError: "Connection error occurred",
    SaveError: "Error occurred while saving",
    LoadError: "Error occurred while loading data",

    // User Friendly Messages
    NoAlertsMessage: "No alerts are currently available",
    AlertsLoadingMessage: "Loading alerts...",
    AlertCreatedSuccess: "Alert created successfully",
    AlertUpdatedSuccess: "Alert updated successfully",
    AlertDeletedSuccess: "Alert deleted successfully",
    SettingsSavedSuccess: "Settings saved successfully",

    // Date and Time
    CreatedBy: "Created by",
    CreatedOn: "Created on",
    LastModified: "Last modified",
    Never: "Never",
    Today: "Today",
    Yesterday: "Yesterday",
    Tomorrow: "Tomorrow",

    // Validation Messages
    FieldRequired: "This field is required",
    InvalidUrl: "Please enter a valid URL",
    InvalidDate: "Please enter a valid date",
    InvalidEmail: "Please enter a valid email address",
    TitleTooLong: "Title is too long (maximum 255 characters)",
    DescriptionTooLong: "Description is too long (maximum 2000 characters)",

    // Rich Media
    UploadImage: "Upload Image",
    RemoveImage: "Remove Image",
    ImageAltText: "Image Alt Text",
    VideoUrl: "Video URL",
    EmbedCode: "Embed Code",

    // Accessibility
    CloseDialog: "Close dialog",
    OpenSettings: "Open settings",
    ExpandAlert: "Expand alert",
    CollapseAlert: "Collapse alert",
    AlertContentLabel: "Alert details",
    AlertActions: "Alert actions",
    PinAlert: "Pin alert",
    UnpinAlert: "Unpin alert",

    // Alert UI
    AlertHeaderPriorityTooltip: "Priority: {0}",
    AttachmentsHeader: "📎 Attachments ({0})",
    FileSizeKilobytes: "{0} KB",
    FileSizeMegabytes: "{0} MB",
    AlertActionsOpenLink: "Open link",
    AlertActionsPrevious: "Previous alert",
    AlertActionsNext: "Next alert",
    AlertActionsDismiss: "Dismiss alert",
    AlertActionsHideForever: "Hide alert forever",
    AlertCarouselCounter: "{0} of {1}",
    DefaultAlertTypeName: "Default",
    AlertsLoadErrorTitle: "Unable to load alerts",
    AlertsLoadErrorFallback:
      "An error occurred while loading alerts. Please try refreshing the page or contact your administrator if the problem persists.",

    // List Management
    FailedToLoadSiteInformation: "Failed to load site information",
    AlertsListCreatedSuccessfully: "Alerts list created successfully on {0}",
    FailedToCreateAlertsList: "Failed to create alerts list",
    HomeSite: "Home Site",
    HubSite: "Hub Site",
    CurrentSite: "Current Site",
    HomeSiteDescription: "Alerts shown on all sites in the tenant",
    HubSiteDescription: "Alerts shown on hub and connected sites",
    CurrentSiteDescription: "Alerts shown only on this site",
    ListExistsAndAccessible: "List exists and accessible",
    ListExistsNoAccess: "List exists but no access",
    ListNotExistsCanCreate: "List not found - can create",
    ListNotExistsCannotCreate: "List not found - cannot create",
    AlertListsManagement: "Alert Lists Management",
    LoadingSiteInformation: "Loading site information...",
    ManageAlertListsDescription:
      "Manage alert lists across your site hierarchy",
    CreatingList: "Creating...",
    CreateAlertsList: "Create Alerts List",
    LanguagesUpdatedSuccessfully: "Languages updated successfully for {0}",
    FailedToUpdateLanguages: "Failed to update languages for {0}",
    SelectLanguagesForList: "Select Languages for Alert List",
    SelectLanguagesDescription:
      "Choose which languages to support for alerts on {0}. English is required and will always be included.",
    SelectedLanguages: "Selected languages",
    CreateWithSelectedLanguages: "Create with {0} languages",
    ManageLanguagesForList: "Manage Languages for Alert List",
    ManageLanguagesDescription:
      "Manage which languages are supported for alerts on {0}. English is required and will always be included.",
    UpdateLanguages: "Update Languages",
    ViewEditLanguages: "Languages",
    ReadyForAlerts: "Ready for alerts",
    AlertHierarchy: "Alert Display Hierarchy",
    AlertHierarchyDescription:
      "Alerts are displayed based on site hierarchy: Home Site alerts appear everywhere, Hub Site alerts appear on hub and connected sites, and Site alerts appear only on the specific site.",
    HomeSiteScope: "Shown on all sites",
    HubSiteScope: "Shown on hub and connected sites",
    SiteScope: "Shown only on this site",

    // Multi-Language Editor
    FailedToLoadLanguages: "Failed to load available languages",
    LanguageAddedSuccessfully: "Language added successfully",
    FailedToAddLanguage: "Failed to add language",
    LanguageRemovedSuccessfully: "Language removed successfully",
    FailedToRemoveLanguage: "Failed to remove language",
    LoadingLanguages: "Loading languages...",
    MultiLanguageContent: "Multi-Language Content",
    MultiLanguageContentDescription:
      "Create content in multiple languages for broader accessibility",
    AddLanguage: "Add Language",
    AddCustomLanguage: "Add Custom Language",
    LanguageCode: "Language Code",
    LanguageName: "Language Name",
    NativeLanguageName: "Native Language Name",
    QuickAdd: "Quick Add",
    SelectSuggestedLanguage: "Select from suggested languages",
    RightToLeft: "Right-to-Left",
    RTLEnabled: "RTL Enabled",
    RTLDisabled: "RTL Disabled",
    EditingContent: "Editing Content",
    RemoveLanguage: "Remove Language",
    ContentSummary: "Content Summary",
    HasContent: "Has Content",
    NoContent: "No Content",

    // Notification Messages
    SettingsUpdatedSuccessfully: "Settings updated successfully!",
    AlertCreatedSuccessfully: "Alert created successfully!",
    AlertUpdatedSuccessfully: "Alert updated successfully!",
    AlertDeletedSuccessfully: "Alert deleted successfully!",
    AlertTypeCreatedSuccessfully: "Alert type created successfully!",
    AlertTypeDeletedSuccessfully: "Alert type deleted successfully!",
    ListsInitializedSuccessfully: "SharePoint lists initialized successfully!",

    // Error Messages
    FailedToCreateAlert:
      "Failed to create alert. Please check your permissions and try again.",
    FailedToUpdateAlert: "Failed to update alert. Please try again.",
    FailedToDeleteAlert: "Failed to delete alert. Please try again.",
    FailedToLoadAlerts: "Failed to load alerts. Please check your permissions.",
    FailedToSaveSettings: "Failed to save settings. Please try again.",
    PermissionDeniedListCreation:
      "You don't have permission to create SharePoint lists. Please contact your administrator.",
    SharePointListsNotFound:
      "SharePoint lists not found and cannot be created automatically.",
    InvalidFormData: "Please fill in all required fields correctly.",
    AlertTypeNameRequired: "Alert type name is required.",
    AlertTypeAlreadyExists: "An alert type with this name already exists.",

    // Form Validation
    TitleRequired: "Alert title is required",
    DescriptionRequired: "Alert description is required",
    AlertTypeRequired: "Please select an alert type",
    TitleMinLength: "Title must be at least 3 characters",
    TitleMaxLength: "Title must be less than 100 characters",
    DescriptionMinLength:
      "Description must be at least 10 characters (excluding HTML tags)",
    InvalidUrlFormat: "Please enter a valid URL",
    LinkDescriptionRequired:
      "Link description is required when URL is provided",
    EndDateMustBeAfterStartDate: "End date must be after start date",
    AtLeastOneSiteRequired:
      "Please select at least one site for alert distribution",

    // Language Field Manager
    MultiLanguageFieldManagement: "Multi-Language Field Management",
    SelectLanguagesToAdd:
      "Select languages to add multi-language content fields to your alert lists. Fields will be created for Title, Description, and Link Description in each selected language.",
    AvailableLanguages: "Available Languages",
    Active: "Active",
    Updating: "Updating...",
    Refresh: "Refresh",
    LoadingLanguageSupport: "Loading language support...",
    LanguageFieldsCreatedAs:
      "Language fields are created as: Title_{LANG}, Description_{LANG}, LinkDescription_{LANG}",
    EnglishCannotBeRemoved:
      "English is the default language and cannot be removed.",
    LanguageSupportAdded: "Added {0} language support successfully!",
    LanguageSupportRemoved: "Removed {0} from active languages.",
    FailedToAddLanguageSupport: "Failed to add {0} language support.",
    FailedToRemoveLanguageSupport: "Failed to remove {0} language support.",
    CouldNotLoadLanguageSupport:
      "Could not load current language support. Using defaults.",

    // Site Selector
    SiteSelectorPermissionOwner: "Owner",
    SiteSelectorPermissionFullControl: "Full Control",
    SiteSelectorPermissionCanEdit: "Can Edit",
    SiteSelectorPermissionDesigner: "Designer",
    SiteSelectorPermissionReadOnly: "Read Only",
    SiteSelectorLoading: "Loading available sites...",
    SiteSelectorQuickSelectionTitle: "Quick Selection",
    SiteSelectorCurrentSiteOnly: "Current Site Only",
    SiteSelectorHubScope: "This Hub ({0} sites)",
    SiteSelectorOrganizationHome: "Organization Home",
    SiteSelectorRecentSites: "Recent Sites ({0})",
    SiteSelectorSearchPlaceholder: "Search sites by name or URL...",
    SiteSelectorFiltersLabel: "Filters",
    SiteSelectorShowWritable: "Show only sites where I can create alerts",
    SiteSelectorSiteTypeLabel: "Site Type:",
    SiteSelectorSiteTypeAll: "All Sites",
    SiteSelectorSiteTypeHub: "Hub Sites",
    SiteSelectorSiteTypeTeam: "Team Sites",
    SiteSelectorSiteTypeCommunication: "Communication Sites",
    SiteSelectorSiteTypeHome: "Home Sites",
    SiteSelectorNoResultsTitle: "No sites found",
    SiteSelectorNoResultsDescription:
      "Try adjusting your search terms or filters.",
    SiteSelectorWritableFilterHint:
      "You may need to disable the 'writable only' filter to see more sites.",
    SiteSelectorSiteSingular: "site",
    SiteSelectorSitePlural: "sites",
    SiteSelectorSelected: "selected",
    SiteSelectorMaxSelection: "(max {0})",

    // Multi-Language Content Editor
    MultiLanguageEditorAddLanguagesLabel: "Add Languages:",
    MultiLanguageEditorAllLanguagesAdded:
      "All available languages have been added",
    MultiLanguageEditorIncompleteBadge: "Incomplete",
    MultiLanguageEditorRemoveButton: "Remove",
    MultiLanguageEditorTitleLabel: "Title",
    MultiLanguageEditorTitlePlaceholder: "Enter alert title in {0}...",
    MultiLanguageEditorDescriptionLabel: "Description",
    MultiLanguageEditorDescriptionPlaceholder:
      "Enter alert description in {0}...",
    MultiLanguageEditorLinkDescriptionLabel: "Link Description",
    MultiLanguageEditorLinkDescriptionPlaceholder:
      "Enter link description in {0}...",
    MultiLanguageEditorFallbackLabel:
      "Use this language as fallback for other users",
    MultiLanguageEditorNoLanguagesTitle: "No Languages Selected",
    MultiLanguageEditorNoLanguagesDescription:
      "Add languages above to create multi-language content",
    MultiLanguageEditorSummaryTitle: "Content Summary:",
    MultiLanguageEditorSummaryComplete: "✓ Complete",
    MultiLanguageEditorSummaryIncomplete: "⚠ Incomplete",
    MultiLanguageEditorLanguageCount: "{0} languages",
    MultiLanguageEditorTranslateAllMissing: "Translate all missing",
    MultiLanguageEditorNoMissingTranslations:
      "All language variants already have title and description.",
    TranslationStatusLabel: "Translation Status",
    TranslationStatusDescription: "Track the state of this language variant",
    TranslationStatusDraft: "Draft",
    TranslationStatusInReview: "In Review",
    TranslationStatusApproved: "Approved",
    LanguagePolicyInheritanceHint: "Empty fields will inherit from {0}.",

    // Language Policy
    LanguagePolicyTitle: "Language Policy",
    LanguagePolicyDescription:
      "Configure translation rules, fallback behavior, and workflow requirements for multi-language alerts.",
    LanguagePolicyFallbackLanguageLabel: "Fallback Language",
    LanguagePolicyFallbackLanguageDescription:
      "Used when a user’s preferred language is not available.",
    LanguagePolicyFallbackTenantDefault: "Tenant default language",
    LanguagePolicyCompletenessRuleLabel: "Translation Completeness",
    LanguagePolicyCompletenessRuleDescription:
      "Choose how complete translations must be before saving or publishing.",
    LanguagePolicyCompletenessAll:
      "Require all selected languages to be complete",
    LanguagePolicyCompletenessAtLeastOne:
      "Require at least one complete language",
    LanguagePolicyCompletenessDefault:
      "Require the default language to be complete",
    LanguagePolicyRequireLinkDescriptionLabel:
      "Require link descriptions when a URL is provided",
    LanguagePolicyRequireLinkDescriptionDescription:
      "Ensures localized link descriptions are provided for each language when links are used.",
    LanguagePolicyPreventDuplicatesLabel: "Prevent duplicate language variants",
    LanguagePolicyPreventDuplicatesDescription:
      "Block saving when the same language is added more than once.",
    LanguagePolicyInheritanceEnable: "Enable field inheritance",
    LanguagePolicyInheritanceDescription:
      "Allow missing fields to inherit from the fallback language.",
    LanguagePolicyInheritanceTitleField: "Inherit Title",
    LanguagePolicyInheritanceDescriptionField: "Inherit Description",
    LanguagePolicyInheritanceLinkDescriptionField: "Inherit Link Description",
    LanguagePolicyWorkflowEnable: "Enable translation workflow",
    LanguagePolicyWorkflowDescription:
      "Track translation status for each language variant.",
    LanguagePolicyWorkflowDefaultStatusLabel: "Default Translation Status",
    LanguagePolicyWorkflowRequireApproved: "Require Approved status to display",
    LanguagePolicyWorkflowRequireApprovedDescription:
      "Hide translation variants that are not approved.",
    LanguagePolicySaveButton: "Save Language Policy",
    LanguagePolicySaving: "Saving...",
    LanguagePolicySavedSuccess: "Language policy saved",
    LanguagePolicySavedFailed: "Failed to save language policy",

    // Language Manager
    LanguageManagerTitle: "Multi-Language Field Management",
    LanguageManagerDescription:
      "Select languages to add multi-language content fields to your alert lists. Fields will be created for Title, Description, and Link Description in each selected language.",
    LanguageManagerAvailableLanguages: "Available Languages",
    LanguageManagerActiveCount: "{0} active",
    LanguageManagerActiveCountWithPending: "{0} active, {1} updating",
    LanguageManagerLoading: "Loading...",
    LanguageManagerLoadingSupport: "Loading language support...",
    LanguageManagerUpdating: "Updating...",
    LanguageManagerActive: "Active",
    LanguageManagerLoadFailedDefault:
      "Could not load language support. Using site default: {0}",
    LanguageManagerDefaultLanguageProtected:
      "{0} is the site's default language and cannot be removed.",
    LanguageManagerAddedSuccess: "Added {0} language support successfully!",
    LanguageManagerRemovedSuccess: "Removed {0} language support and columns.",
    LanguageManagerUpdateFailed: "Failed to update language support for {0}.",

    // Alert Templates
    AlertTemplatesTemplateDescription: 'Template based on "{0}"',
    AlertTemplatesCategoryAll: "All Templates",
    AlertTemplatesCategoryMaintenance: "Maintenance",
    AlertTemplatesCategoryInformation: "Information",
    AlertTemplatesCategoryEmergency: "Emergency",
    AlertTemplatesCategoryInterruption: "Interruption",
    AlertTemplatesCategoryTraining: "Training",
    AlertTemplatesCategoryAnnouncements: "Announcements",
    AlertTemplatesHeaderTitle: "Choose a Template",
    AlertTemplatesHeaderDescription:
      "Start with a pre-configured template and customize it to your needs",
    AlertTemplatesSearchPlaceholder: "Search templates...",
    AlertTemplatesUseTemplate: "Use Template →",
    AlertTemplatesLoadingTitle: "Loading Templates...",
    AlertTemplatesLoadingDescription: "Fetching templates from SharePoint",
    AlertTemplatesNoResultsTitle: "No templates found",
    AlertTemplatesNoResultsDescription:
      "Try adjusting your search terms or category filter",
    AlertTemplatesEmptyTitle: "No templates available",
    AlertTemplatesEmptyDescription:
      "Templates will be created automatically when alert lists are set up",
    AlertTemplatesPinnedBadge: "PINNED",
    AlertTemplatesNotifyBadge: "NOTIFY",
    AlertTemplatesDefaultLinkDescription: "Learn more",

    // Attachments
    AttachmentManagerDeleteConfirm: 'Are you sure you want to delete "{0}"?',
    AttachmentManagerTitle: "Attachments",
    AttachmentManagerUploadingLabel: "Uploading...",
    AttachmentManagerAddFilesButton: "Add Files",
    AttachmentManagerSaveNotice: "Save the alert first to add attachments",
    AttachmentManagerDropZoneInstruction:
      'Drag and drop files here or click "Add Files" button',
    AttachmentManagerDropZoneHint: "{0} | Max {1}MB per file",
    AttachmentManagerProgressTitle: "Uploading files...",
    AttachmentManagerStatusCompleted: "✓ Completed",
    AttachmentManagerStatusFailed: "✗ Failed",
    AttachmentManagerDownloadTitle: "Download",
    AttachmentManagerDeleteTitle: "Delete",
    AttachmentManagerHelpText:
      "Allowed formats: {0} | Max size: {1}MB per file",
    AttachmentManagerDeleteError:
      "Failed to delete attachment. Please try again.",
    AttachmentManagerSaveAlertFirst:
      "Please save the alert first before adding attachments.",
    AttachmentManagerUploadFailed:
      "Failed to upload some attachments. Please try again.",
    AttachmentManagerInvalidFormat: "{0}: Invalid file format (allowed: {1})",
    AttachmentManagerFileTooLarge: "{0}: File size exceeds {1}MB limit ({2}MB)",

    // Image Manager
    ImageManagerDeleteConfirm: 'Delete "{0}"? This cannot be undone.',
    ImageManagerLoading: "Loading images...",
    ImageManagerError: "Error: {0}",
    ImageManagerEmpty: "No images uploaded yet",
    ImageManagerHeaderTitle: "Uploaded Images ({0})",
    ImageManagerCopyUrl: "Copy URL",
    ImageManagerDeleteError: "Failed to delete image: {0}",
    ImageManagerDeleteSuccess: "Image deleted successfully.",
    ImageManagerCopySuccess: "Image URL copied to clipboard!",
    ImageManagerCopyError: "Failed to copy URL",
    ImageUploadFailure: "Image upload failed. Please try again.",
    ImageUploadInvalidFile:
      "Please select a valid image file (PNG, JPG, GIF, or WebP).",

    // Repair Dialog
    RepairDialogWarningMessage:
      "This action will modify your SharePoint list structure.",
    RepairDialogActionsIntro: "The repair process will:",
    RepairDialogRemoveColumns:
      "Remove outdated columns that are no longer needed",
    RepairDialogAddColumns: "Add missing columns with current definitions",
    RepairDialogUpdateColumns:
      "Update existing columns to match the latest schema",
    RepairDialogProtectData:
      "Preserve all existing data and language-specific columns",
    RepairDialogSafeProcess: "Safe Process:",
    RepairDialogSafeProcessDescription:
      "This repair is designed to be non-destructive. Your existing alert data will be preserved, and language-specific columns will be kept intact.",
    RepairDialogProgressTitle: "Repairing Alerts List...",
    RepairDialogProgressDescription:
      "Please wait while we update your list structure. This may take a few minutes.",
    RepairDialogResultSuccessTitle: "Repair Completed!",
    RepairDialogResultFailureTitle: "Repair Failed",
    RepairDialogRemovedColumnsTitle: "Removed Columns ({0})",
    RepairDialogAddedColumnsTitle: "Added/Updated Columns ({0})",
    RepairDialogWarningsTitle: "Warnings ({0})",
    RepairDialogErrorsTitle: "Errors ({0})",
    RepairDialogStartButton: "Start Repair",
    RepairDialogCancelButton: "Cancel",
    RepairDialogInProgressButton: "Repairing...",
    RepairDialogCloseButton: "Close",
    RepairDialogDialogTitleSuccess: "Repair Completed Successfully",
    RepairDialogDialogTitleIssues: "Repair Completed with Issues",
    RepairDialogDialogTitleProgress: "Repairing Alerts List",
    RepairDialogTitle: "Repair Alerts List",
    RepairDialogPleaseWait: "Please wait...",

    // Color Picker
    ColorPickerSelectedColorAria: "Selected color: {0}",
    ColorPickerPresetColorsTitle: "Preset Colors",
    ColorPickerCustomColorTitle: "Custom Color",
    ColorPickerApplyButton: "Apply",
    ColorPickerSelectColorAria: "Select color {0}",

    // Emoji Picker
    EmojiPickerButtonTitle: "Add emoji",
    EmojiPickerButtonLabel: "Add Emoji",
    EmojiPickerSearchPlaceholder: "Search emoji...",

    // Rich Text Editor
    RichTextEditorLoadError: "Unable to load Rich Text editor",
    RichTextEditorPlaceholder: "Enter your message...",

    // Progress Indicator
    ProgressIndicatorTitle: "Progress",
    ProgressIndicatorFailedSummary: "{0} of {1} steps failed",
    ProgressIndicatorCompletedSummary: "{0} of {1} steps completed",
    ProgressIndicatorError: "Error: {0}",

    // Alert Preview
    AlertPreviewDefaultTitle: "Alert Title",
    AlertPreviewDefaultDescription: "Alert description will appear here...",
    AlertPreviewPinnedBadge: "📌 PINNED",
    AlertPreviewTypeLabel: "Type:",
    AlertPreviewPriorityLabel: "Priority:",
    AlertPreviewBackgroundLabel: "Background:",
    AlertPreviewTextColorLabel: "Text Color:",
    AlertPreviewIconLabel: "Icon:",
    AlertPreviewLinkPrefix: "🔗 {0}",

    // Error Boundary
    ErrorBoundaryHeader: "Something went wrong in {0}",
    ErrorBoundaryFallbackMessage: "An unexpected error occurred",
    ErrorBoundaryIdLabel: "Error ID: {0}",
    ErrorBoundaryRetryButton: "Try Again",
    ErrorBoundaryAttemptsLeft: "attempts left",
    ErrorBoundaryCopyButton: "Copy Error Details",
    ErrorBoundaryReloadButton: "Reload Page",
    ErrorBoundaryDevDetailsTitle: "Development Error Details",

    // Permission Status
    PermissionStatusChecking: "Checking permissions...",
    PermissionStatusAllGranted: "All required permissions are granted",
    PermissionStatusMissingTitle: "Additional permissions required",
    PermissionStatusSitesWriteRequired: "Sites.ReadWrite.All",
    PermissionStatusSitesWriteDescription:
      "Required to create and manage alert lists across sites",
    PermissionStatusMailRequired: "Mail.Send",
    PermissionStatusMailDescription:
      "Required to send email notifications for critical alerts",
    PermissionStatusShowDetails: "Show details",
    PermissionStatusHideDetails: "Hide details",
    PermissionStatusMissingList: "Missing permissions",
    PermissionStatusAdminAction:
      "As a site admin, you can grant these permissions:",
    PermissionStatusGrantPermissions: "Grant permissions in Azure AD",
    PermissionStatusContactAdmin:
      "Please contact your SharePoint administrator to grant these permissions.",
    ManageAlertsAlertTypeFromSourceSiteLabel: "{0} (from source site)",


    ManageAlertsCrossSiteImagesNotice:
      "Note: Images from other sites may not display correctly due to cross-site restrictions.",

    // Copilot Features
    EnableCopilotLabel: "Enable Copilot Assistance (Preview)",
    EnableCopilotDescription:
      "Use Microsoft Copilot to draft alerts, translate content, and analyze content tone and professionalism. Requires Copilot for Microsoft 365 license.",
    CopilotOverwriteConfirmation:
      "This will overwrite existing content. Continue?",
    CopilotDefaultLanguageRequired:
      "Please fill in the default language content first.",
    CopilotTranslationFailed: "Translation failed. Please try again.",
    CopilotDraftButton: "Draft with Copilot",
    CopilotDraftTitle: "Draft with Copilot",
    CopilotPreviewTitle: "Preview",
    CopilotAcceptDraft: "Use This Draft",
    CopilotKeywordsLabel: "What should the alert say?",
    CopilotKeywordsPlaceholder:
      "e.g., Warning about email server maintenance tonight...",
    CopilotToneLabel: "Tone",
    CopilotToneProfessional: "Professional (Default)",
    CopilotToneUrgent: "Urgent",
    CopilotToneCasual: "Casual",
    CopilotGeneratingLabel: "Generating...",
    CopilotGenerateButton: "Generate Draft",
    CopilotDraftGenerationFailed: "Failed to generate draft",
    CopilotUnexpectedError: "An unexpected error occurred",
    CopilotErrorTitle: "Copilot Error",
    CopilotCheckingLabel: "Checking...",
    CopilotSentimentButton: "Sentiment Check",
    CopilotSentimentAnalysisTitle: "Sentiment Analysis Result",
    CopilotNoIssuesFound: "No issues found.",
    CopilotAnalyzeFailed: "Failed to analyze text",
    CopilotTranslatingLabel: "Translating...",
    CopilotTranslateButton: "Translate with Copilot",
    CopilotOverwriteConfirmationTitle: "Overwrite Translation",
    CopilotOverwriteConfirmButton: "Overwrite",
    CopilotSentimentStatusGreen: "Looks Good",
    CopilotSentimentStatusYellow: "Review Suggested",
    CopilotSentimentStatusRed: "Issues Found",
    CopilotSentimentSummaryLabel: "Full Analysis",

    // Copilot Draft - Tips and Examples
    CopilotTipsToggleLabel: "How to write a good prompt",
    CopilotTipsCloseAriaLabel: "Close tips",
    CopilotTipsItem1: "Be specific about the action or announcement",
    CopilotTipsItem2: "Mention dates, times, and deadlines clearly",
    CopilotTipsItem3: "Include who the alert affects",
    CopilotTipsItem4: "Add any required actions for readers",
    CopilotExamplesLabel: "Try an example:",
    CopilotExamplePrompt1: "Announce new security policy requiring 2FA",
    CopilotExamplePrompt2: "System maintenance tonight at 10 PM",
    CopilotExamplePrompt3: "Welcome new team members to the department",
    CopilotExamplePrompt4: "Office closure due to severe weather",
    CopilotInputLabel: "What should the alert say?",
    CopilotInputPlaceholder: "Describe your announcement, maintenance window, policy change, etc.",
    CopilotCharLimitError: "Please keep your prompt under 500 characters",
    CopilotKeyboardHint: "Press Ctrl+Enter to generate",
    CopilotToneSelectorLabel: "Choose a tone:",
    CopilotToneProfessionalShort: "Professional",
    CopilotToneUrgentShort: "Urgent",
    CopilotToneCasualShort: "Casual",
    CopilotGeneratingSubtext: "Crafting your {0} alert...",
    CopilotPreviewBadge: "{0}",
    CopilotDraftStatsWords: "{0} words",
    CopilotDraftStatsChars: "{0} characters",
    CopilotRefineSectionLabel: "Refine the draft:",
    CopilotRefineShorter: "Shorter",
    CopilotRefineLonger: "Longer",
    CopilotRefineRephrase: "Rephrase",
    CopilotRefineTryAgain: "Try Again",
    CopilotRefiningLabel: "Refining...",
    CopilotEditButton: "Edit",

    // Copilot Sentiment - Fix Actions
    CopilotSentimentFixTrim: "Trim to 50 words",
    CopilotSentimentFixConcise: "Make concise",
    CopilotSentimentTooltipDisabled: "Add some text to your description first",
    CopilotSentimentTooltipExpand: "Click to view details",
    CopilotSentimentTooltipMinimize: "Minimize",
    CopilotSentimentTooltipDismiss: "Dismiss",
    CopilotSentimentDismissAriaLabel: "Dismiss analysis",
    CopilotSentimentReadingStatsWords: "words",
    CopilotSentimentReadingStatsTime: "read time",
    CopilotSentimentToneLabel: "tone",
    CopilotSentimentIssuesFixedLabel: "Issues Fixed",
    CopilotSentimentIssuesFixedCount: "{0} of {1}",
    CopilotSentimentCriticalSection: "Critical",
    CopilotSentimentRecommendedSection: "Recommended",
    CopilotSentimentOptionalSection: "Suggestions",
    CopilotSentimentAnalysisProfessional: "Professional",
    CopilotSentimentAnalysisTone: "Tone",

    // About Section
    SettingsAboutTitle: "About",
    SettingsAboutProductName: "Alert Banner",
    SettingsAboutVersion: "Version {0}",
    SettingsAboutDescription: "A free, open-source SharePoint SPFx extension for managing alert banners across your tenant.",
    SettingsAboutAuthor: "Created with ❤️ by the community",
    SettingsAboutGitHubLink: "View on GitHub",
    SettingsAboutGitHubUrl: "https://github.com/nicolaskheirallah/AlertBanner",
  };
});
