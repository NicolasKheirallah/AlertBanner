import {
  TargetLanguage,
  IAlertItem,
  ILanguageContent,
} from "../Alerts/IAlerts";
import { validationService } from "../Services/ValidationService";
import { VALIDATION_MESSAGES } from "./AppConstants";
import {
  ILanguagePolicy,
  normalizeLanguagePolicy,
} from "../Services/LanguagePolicyService";

export interface IFormErrors {
  [key: string]: string;
}

export interface IValidationOptions {
  useMultiLanguage: boolean;
  getString?: (key: string, ...args: any[]) => string;
  languagePolicy?: ILanguagePolicy;
  tenantDefaultLanguage?: TargetLanguage;
  validateTargetSites?: boolean;
}

export const validateAlertData = (
  alert: Partial<IAlertItem>,
  options: IValidationOptions,
): IFormErrors => {
  const { useMultiLanguage, getString } = options;
  const errors: IFormErrors = {};

  // Helper function to get localized string or fallback
  const getLocalizedString = (key: string, ...args: any[]): string => {
    if (getString) {
      return getString(key, ...args);
    }
    // Fallback messages if localization is not available
    const fallbackMessages: { [key: string]: string } = VALIDATION_MESSAGES;

    let message = fallbackMessages[key] || key;
    // Simple placeholder replacement
    args.forEach((arg, index) => {
      message = message.replace(`{${index}}`, arg);
    });
    return message;
  };

  // Multi-language validation
  if (useMultiLanguage) {
    const policy = normalizeLanguagePolicy(options.languagePolicy);
    const languageContent = alert.languageContent as
      | ILanguageContent[]
      | undefined;

    if (!languageContent || languageContent.length === 0) {
      errors.title = getLocalizedString("CreateAlertLanguageRequired");
    } else {
      const fallbackLanguage =
        policy.fallbackLanguage === "tenant-default"
          ? options.tenantDefaultLanguage || TargetLanguage.EnglishUS
          : policy.fallbackLanguage;
      const fallbackContent = languageContent.find(
        (content) => content.language === fallbackLanguage,
      );
      const resolveFieldSatisfied = (
        content: ILanguageContent,
        field: "title" | "description" | "linkDescription",
      ): boolean => {
        const value = content[field];
        if (value && value.trim().length > 0) {
          return true;
        }
        if (!policy.inheritance.enabled) {
          return false;
        }
        if (!policy.inheritance.fields[field]) {
          return false;
        }
        const fallbackValue = fallbackContent
          ? fallbackContent[field]
          : undefined;
        return !!fallbackValue && fallbackValue.trim().length > 0;
      };

      // Duplicate language check removed

      let hasCompleteLanguage = false;

      languageContent.forEach((content, index) => {
        const languagePrefix = `${content.language}_${index}`;
        const titleTrimmed = content.title?.trim() || "";
        const descriptionTrimmed = content.description?.trim() || "";
        const linkTrimmed = content.linkDescription?.trim() || "";
        const hasAnyContent =
          !!titleTrimmed || !!descriptionTrimmed || !!linkTrimmed;

        const titleOk =
          (resolveFieldSatisfied(content, "title") &&
            titleTrimmed.length >= 3) ||
          (!titleTrimmed && resolveFieldSatisfied(content, "title"));
        const descriptionOk =
          (resolveFieldSatisfied(content, "description") &&
            descriptionTrimmed.length >= 10) ||
          (!descriptionTrimmed &&
            resolveFieldSatisfied(content, "description"));
        const linkRequired =
          policy.requireLinkDescriptionWhenUrl && !!alert.linkUrl;
        const linkOk =
          !linkRequired || resolveFieldSatisfied(content, "linkDescription");

        // Validate title length
        const enforceThisLanguage =
          policy.completenessRule === "allSelectedComplete" ||
          (policy.completenessRule === "requireDefaultLanguageComplete" &&
            content.language === fallbackLanguage) ||
          (policy.completenessRule === "atLeastOneComplete" && hasAnyContent);

        if (enforceThisLanguage) {
          if (!titleTrimmed && !resolveFieldSatisfied(content, "title")) {
            errors[`title_${languagePrefix}`] = getLocalizedString(
              "CreateAlertLanguageTitleRequired",
              content.language,
            );
            if (index === 0 && !errors.title) {
              errors.title = getLocalizedString(
                "CreateAlertLanguageTitleRequired",
                content.language,
              );
            }
          } else if (titleTrimmed && titleTrimmed.length < 3) {
            errors[`title_${languagePrefix}`] =
              getLocalizedString("TitleMinLength");
            if (index === 0 && !errors.title) {
              errors.title = getLocalizedString("TitleMinLength");
            }
          }

          if (
            !descriptionTrimmed &&
            !resolveFieldSatisfied(content, "description")
          ) {
            errors[`description_${languagePrefix}`] = getLocalizedString(
              "CreateAlertLanguageDescriptionRequired",
              content.language,
            );
            if (index === 0 && !errors.description) {
              errors.description = getLocalizedString(
                "CreateAlertLanguageDescriptionRequired",
                content.language,
              );
            }
          } else if (descriptionTrimmed && descriptionTrimmed.length < 10) {
            errors[`description_${languagePrefix}`] = getLocalizedString(
              "DescriptionMinLength",
            );
            if (index === 0 && !errors.description) {
              errors.description = getLocalizedString("DescriptionMinLength");
            }
          }

          if (linkRequired && !linkOk) {
            errors[`linkDescription_${languagePrefix}`] = getLocalizedString(
              "CreateAlertLanguageLinkDescriptionRequired",
              content.language,
            );
            if (index === 0 && !errors.linkDescription) {
              errors.linkDescription = getLocalizedString(
                "CreateAlertLanguageLinkDescriptionRequired",
                content.language,
              );
            }
          }
        }

        // Check if this language is complete
        if (titleOk && descriptionOk && (!linkRequired || linkOk)) {
          hasCompleteLanguage = true;
        }
      });

      // Ensure at least one language has complete content
      if (
        policy.completenessRule === "atLeastOneComplete" &&
        !hasCompleteLanguage &&
        languageContent.length > 0
      ) {
        errors.languageContent = getLocalizedString(
          "CreateAlertLanguageAtLeastOneComplete",
        );
      }

      if (
        policy.completenessRule === "requireDefaultLanguageComplete" &&
        languageContent.length > 0
      ) {
        const defaultLangContent = languageContent.find(
          (content) => content.language === fallbackLanguage,
        );
        const defaultComplete = defaultLangContent
          ? resolveFieldSatisfied(defaultLangContent, "title") &&
            resolveFieldSatisfied(defaultLangContent, "description") &&
            (!policy.requireLinkDescriptionWhenUrl ||
              !alert.linkUrl ||
              resolveFieldSatisfied(defaultLangContent, "linkDescription"))
          : false;
        if (!defaultComplete) {
          errors.languageContent = getLocalizedString(
            "CreateAlertDefaultLanguageRequired",
            fallbackLanguage,
          );
        }
      }
    }
  } else {
    // Single language validation
    if (!alert.title?.trim()) {
      errors.title = getLocalizedString("TitleRequired");
    } else if (alert.title.trim().length < 3) {
      errors.title = getLocalizedString("TitleMinLength");
    } else if (alert.title.length > 100) {
      errors.title = getLocalizedString("TitleMaxLength");
    }

    if (!alert.description?.trim()) {
      errors.description = getLocalizedString("DescriptionRequired");
    } else if (alert.description.length < 10) {
      errors.description = getLocalizedString("DescriptionMinLength");
    }

    if (alert.linkUrl && !alert.linkDescription?.trim()) {
      errors.linkDescription = getLocalizedString("LinkDescriptionRequired");
    }
  }

  // Common validations (apply regardless of multi-language mode)

  // Alert type validation
  if (!alert.AlertType || !alert.AlertType.trim()) {
    errors.AlertType = getLocalizedString("AlertTypeRequired");
  }

  // URL format validation using ValidationService for comprehensive security checks
  if (alert.linkUrl && alert.linkUrl.trim()) {
    const urlValidation = validationService.validateUrl(alert.linkUrl);
    if (!urlValidation.isValid) {
      // Use the first error from the comprehensive validation
      errors.linkUrl =
        urlValidation.errors[0] || getLocalizedString("InvalidUrlFormat");
    }
  }

  // Target sites validation
  if (
    options.validateTargetSites !== false &&
    (!alert.targetSites || alert.targetSites.length === 0)
  ) {
    errors.targetSites = getLocalizedString("AtLeastOneSiteRequired");
  }

  // Date validation
  if (alert.scheduledStart && alert.scheduledEnd) {
    const startDate =
      typeof alert.scheduledStart === "string"
        ? new Date(alert.scheduledStart)
        : alert.scheduledStart;
    const endDate =
      typeof alert.scheduledEnd === "string"
        ? new Date(alert.scheduledEnd)
        : alert.scheduledEnd;

    if (startDate >= endDate) {
      errors.scheduledEnd = getLocalizedString("EndDateMustBeAfterStartDate");
    }
  }

  return errors;
};

export const hasNoErrors = (errors: IFormErrors): boolean => {
  return Object.keys(errors).length === 0;
};

export const getFieldError = (
  errors: IFormErrors,
  fieldName: string,
): string => {
  return errors[fieldName] || "";
};
