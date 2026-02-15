/**
 * @file SharedTypes.ts
 * @description Shared type definitions used by multiple tab components
 *   (CreateAlertTab, ManageAlertsTab) to avoid interface duplication.
 */

/**
 * Form validation errors for alert create/edit forms.
 * Keys correspond to alert field names; values are human-readable error strings.
 * The index signature supports dynamic language-specific error keys.
 */
export interface IFormErrors {
  title?: string;
  description?: string;
  AlertType?: string;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  /** Index signature for dynamic language error keys */
  [key: string]: string | undefined;
}
