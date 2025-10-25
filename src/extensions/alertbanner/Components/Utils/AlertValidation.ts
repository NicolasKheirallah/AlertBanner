import { IAlertItem } from "../Services/SharePointAlertService";
import { ILanguageContent } from "../Services/LanguageAwarenessService";
import { validationService } from "../Services/ValidationService";

export interface IFormErrors {
  [key: string]: string;
}

export interface IValidationOptions {
  useMultiLanguage: boolean;
  getString?: (key: string, ...args: any[]) => string;
}

/**
 * Validates alert data for both create and edit operations
 * @param alert - The alert data to validate (partial for creation, full for editing)
 * @param options - Validation options including multi-language flag and localization function
 * @returns Object containing validation errors, empty object if no errors
 */
export const validateAlertData = (
  alert: any, // Use any to support both IAlertItem and INewAlert/IEditingAlert with Date types
  options: IValidationOptions
): IFormErrors => {
  const { useMultiLanguage, getString } = options;
  const errors: IFormErrors = {};

  // Helper function to get localized string or fallback
  const getLocalizedString = (key: string, ...args: any[]): string => {
    if (getString) {
      return getString(key, ...args);
    }
    // Fallback messages if localization is not available
    const fallbackMessages: { [key: string]: string } = {
      'CreateAlertLanguageRequired': 'At least one language must be configured',
      'CreateAlertLanguageTitleRequired': 'Title is required for {0}',
      'CreateAlertLanguageDescriptionRequired': 'Description is required for {0}',
      'CreateAlertLanguageLinkDescriptionRequired': 'Link description is required for {0} when URL is provided',
      'TitleRequired': 'Title is required',
      'TitleMinLength': 'Title must be at least 3 characters',
      'TitleMaxLength': 'Title cannot exceed 100 characters',
      'DescriptionRequired': 'Description is required',
      'DescriptionMinLength': 'Description must be at least 10 characters',
      'LinkDescriptionRequired': 'Link description is required when URL is provided',
      'AlertTypeRequired': 'Alert type is required',
      'InvalidUrlFormat': 'Please enter a valid URL',
      'AtLeastOneSiteRequired': 'At least one target site must be selected',
      'EndDateMustBeAfterStartDate': 'End date must be after start date',
    };

    let message = fallbackMessages[key] || key;
    // Simple placeholder replacement
    args.forEach((arg, index) => {
      message = message.replace(`{${index}}`, arg);
    });
    return message;
  };

  // Multi-language validation
  if (useMultiLanguage) {
    const languageContent = alert.languageContent as ILanguageContent[] | undefined;

    if (!languageContent || languageContent.length === 0) {
      errors.title = getLocalizedString('CreateAlertLanguageRequired');
    } else {
      languageContent.forEach((content, index) => {
        const languagePrefix = `${content.language}_${index}`;

        if (!content.title?.trim()) {
          errors[`title_${languagePrefix}`] = getLocalizedString('CreateAlertLanguageTitleRequired', content.language);
          // Also set generic title error for first language
          if (index === 0 && !errors.title) {
            errors.title = getLocalizedString('CreateAlertLanguageTitleRequired', content.language);
          }
        }

        if (!content.description?.trim()) {
          errors[`description_${languagePrefix}`] = getLocalizedString('CreateAlertLanguageDescriptionRequired', content.language);
          // Also set generic description error for first language
          if (index === 0 && !errors.description) {
            errors.description = getLocalizedString('CreateAlertLanguageDescriptionRequired', content.language);
          }
        }

        if (alert.linkUrl && !content.linkDescription?.trim()) {
          errors[`linkDescription_${languagePrefix}`] = getLocalizedString('CreateAlertLanguageLinkDescriptionRequired', content.language);
          // Also set generic linkDescription error for first language
          if (index === 0 && !errors.linkDescription) {
            errors.linkDescription = getLocalizedString('CreateAlertLanguageLinkDescriptionRequired', content.language);
          }
        }
      });
    }
  } else {
    // Single language validation
    if (!alert.title?.trim()) {
      errors.title = getLocalizedString('TitleRequired');
    } else if (alert.title.length < 3) {
      errors.title = getLocalizedString('TitleMinLength');
    } else if (alert.title.length > 100) {
      errors.title = getLocalizedString('TitleMaxLength');
    }

    if (!alert.description?.trim()) {
      errors.description = getLocalizedString('DescriptionRequired');
    } else if (alert.description.length < 10) {
      errors.description = getLocalizedString('DescriptionMinLength');
    }

    if (alert.linkUrl && !alert.linkDescription?.trim()) {
      errors.linkDescription = getLocalizedString('LinkDescriptionRequired');
    }
  }

  // Common validations (apply regardless of multi-language mode)

  // Alert type validation
  if (!alert.AlertType || !alert.AlertType.trim()) {
    errors.AlertType = getLocalizedString('AlertTypeRequired');
  }

  // URL format validation using ValidationService for comprehensive security checks
  if (alert.linkUrl && alert.linkUrl.trim()) {
    const urlValidation = validationService.validateUrl(alert.linkUrl);
    if (!urlValidation.isValid) {
      // Use the first error from the comprehensive validation
      errors.linkUrl = urlValidation.errors[0] || getLocalizedString('InvalidUrlFormat');
    }
  }

  // Target sites validation
  if (!alert.targetSites || alert.targetSites.length === 0) {
    errors.targetSites = getLocalizedString('AtLeastOneSiteRequired');
  }

  // Date validation
  if (alert.scheduledStart && alert.scheduledEnd) {
    const startDate = typeof alert.scheduledStart === 'string'
      ? new Date(alert.scheduledStart)
      : alert.scheduledStart;
    const endDate = typeof alert.scheduledEnd === 'string'
      ? new Date(alert.scheduledEnd)
      : alert.scheduledEnd;

    if (startDate >= endDate) {
      errors.scheduledEnd = getLocalizedString('EndDateMustBeAfterStartDate');
    }
  }

  return errors;
};

/**
 * Checks if there are any validation errors
 * @param errors - The errors object returned from validateAlertData
 * @returns true if there are no errors, false otherwise
 */
export const hasNoErrors = (errors: IFormErrors): boolean => {
  return Object.keys(errors).length === 0;
};

/**
 * Gets a user-friendly error message for a specific field
 * @param errors - The errors object
 * @param fieldName - The field name to get error for
 * @returns The error message or empty string if no error
 */
export const getFieldError = (errors: IFormErrors, fieldName: string): string => {
  return errors[fieldName] || '';
};
