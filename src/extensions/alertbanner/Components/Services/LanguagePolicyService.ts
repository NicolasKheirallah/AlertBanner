import { TargetLanguage, TranslationStatus } from "../Alerts/IAlerts";

export type TranslationCompletenessRule =
  | "atLeastOneComplete"
  | "allSelectedComplete"
  | "requireDefaultLanguageComplete";

export interface ILanguagePolicy {
  version: number;
  fallbackLanguage: TargetLanguage | "tenant-default";
  completenessRule: TranslationCompletenessRule;
  requireLinkDescriptionWhenUrl: boolean;
  preventDuplicateLanguages: boolean;
  inheritance: {
    enabled: boolean;
    fields: {
      title: boolean;
      description: boolean;
      linkDescription: boolean;
    };
  };
  workflow: {
    enabled: boolean;
    defaultStatus: TranslationStatus;
    requireApprovedForDisplay: boolean;
  };
}

export const DEFAULT_LANGUAGE_POLICY: ILanguagePolicy = {
  version: 1,
  fallbackLanguage: TargetLanguage.EnglishUS,
  completenessRule: "allSelectedComplete",
  requireLinkDescriptionWhenUrl: true,
  preventDuplicateLanguages: true,
  inheritance: {
    enabled: false,
    fields: {
      title: true,
      description: true,
      linkDescription: true
    }
  },
  workflow: {
    enabled: false,
    defaultStatus: TranslationStatus.Draft,
    requireApprovedForDisplay: true
  }
};

const validCompletenessRules: TranslationCompletenessRule[] = [
  "atLeastOneComplete",
  "allSelectedComplete",
  "requireDefaultLanguageComplete"
];

export const normalizeLanguagePolicy = (policy?: Partial<ILanguagePolicy>): ILanguagePolicy => {
  const merged: ILanguagePolicy = {
    ...DEFAULT_LANGUAGE_POLICY,
    ...policy,
    inheritance: {
      ...DEFAULT_LANGUAGE_POLICY.inheritance,
      ...(policy?.inheritance || {}),
      fields: {
        ...DEFAULT_LANGUAGE_POLICY.inheritance.fields,
        ...(policy?.inheritance?.fields || {})
      }
    },
    workflow: {
      ...DEFAULT_LANGUAGE_POLICY.workflow,
      ...(policy?.workflow || {})
    }
  };

  if (!validCompletenessRules.includes(merged.completenessRule)) {
    merged.completenessRule = DEFAULT_LANGUAGE_POLICY.completenessRule;
  }

  if (!merged.fallbackLanguage) {
    merged.fallbackLanguage = DEFAULT_LANGUAGE_POLICY.fallbackLanguage;
  }

  return merged;
};
