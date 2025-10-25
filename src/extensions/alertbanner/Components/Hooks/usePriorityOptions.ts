import { useMemo } from 'react';
import { AlertPriority } from '../Alerts/IAlerts';

export interface ISharePointSelectOption {
  value: string;
  label: string;
}

/**
 * Custom hook to generate priority options for dropdown/select components
 * @param getString - Optional localization function for getting translated strings
 * @returns Array of formatted priority options
 */
export const usePriorityOptions = (
  getString?: (key: string) => string
): ISharePointSelectOption[] => {
  return useMemo(() => {
    // Helper function to get localized string or fallback
    const getLocalizedLabel = (key: string, fallback: string): string => {
      if (getString) {
        try {
          return getString(key);
        } catch {
          return fallback;
        }
      }
      return fallback;
    };

    return [
      {
        value: AlertPriority.Low,
        label: getLocalizedLabel(
          'CreateAlertPriorityLowDescription',
          'Low Priority - Informational updates and general notices'
        )
      },
      {
        value: AlertPriority.Medium,
        label: getLocalizedLabel(
          'CreateAlertPriorityMediumDescription',
          'Medium Priority - General announcements and updates'
        )
      },
      {
        value: AlertPriority.High,
        label: getLocalizedLabel(
          'CreateAlertPriorityHighDescription',
          'High Priority - Important notices requiring attention'
        )
      },
      {
        value: AlertPriority.Critical,
        label: getLocalizedLabel(
          'CreateAlertPriorityCriticalDescription',
          'Critical Priority - Urgent actions required immediately'
        )
      }
    ];
  }, [getString]);
};
