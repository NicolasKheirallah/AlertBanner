import { Text } from "@microsoft/sp-core-library";
import * as strings from "AlertBannerApplicationCustomizerStrings";

export const VALIDATION_STRING_KEYS = [
  "CreateAlertLanguageRequired",
  "CreateAlertLanguageTitleRequired",
  "CreateAlertLanguageDescriptionRequired",
  "CreateAlertLanguageLinkDescriptionRequired",
  "CreateAlertLanguageAtLeastOneComplete",
  "CreateAlertDefaultLanguageRequired",
  "TitleRequired",
  "TitleMinLength",
  "TitleMaxLength",
  "DescriptionRequired",
  "DescriptionMinLength",
  "LinkDescriptionRequired",
  "AlertTypeRequired",
  "InvalidUrlFormat",
  "AtLeastOneSiteRequired",
  "EndDateMustBeAfterStartDate",
] as const;

type ValidationStringKey = (typeof VALIDATION_STRING_KEYS)[number];

const VALIDATION_MESSAGES: Record<ValidationStringKey, string> = {
  CreateAlertLanguageRequired: strings.CreateAlertLanguageRequired,
  CreateAlertLanguageTitleRequired: strings.CreateAlertLanguageTitleRequired,
  CreateAlertLanguageDescriptionRequired:
    strings.CreateAlertLanguageDescriptionRequired,
  CreateAlertLanguageLinkDescriptionRequired:
    strings.CreateAlertLanguageLinkDescriptionRequired,
  CreateAlertLanguageAtLeastOneComplete:
    strings.CreateAlertLanguageAtLeastOneComplete,
  CreateAlertDefaultLanguageRequired: strings.CreateAlertDefaultLanguageRequired,
  TitleRequired: strings.TitleRequired,
  TitleMinLength: strings.TitleMinLength,
  TitleMaxLength: strings.TitleMaxLength,
  DescriptionRequired: strings.DescriptionRequired,
  DescriptionMinLength: strings.DescriptionMinLength,
  LinkDescriptionRequired: strings.LinkDescriptionRequired,
  AlertTypeRequired: strings.AlertTypeRequired,
  InvalidUrlFormat: strings.InvalidUrlFormat,
  AtLeastOneSiteRequired: strings.AtLeastOneSiteRequired,
  EndDateMustBeAfterStartDate: strings.EndDateMustBeAfterStartDate,
};

const isValidationStringKey = (key: string): key is ValidationStringKey =>
  Object.prototype.hasOwnProperty.call(VALIDATION_MESSAGES, key);

export const getLocalizedValidationMessage = (
  key: string,
  ...args: Array<string | number>
): string => {
  const template = isValidationStringKey(key) ? VALIDATION_MESSAGES[key] : key;
  if (args.length === 0) {
    return template;
  }
  return Text.format(template, ...args.map((arg) => arg.toString()));
};
