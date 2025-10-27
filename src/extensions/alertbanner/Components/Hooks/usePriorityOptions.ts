import { useMemo } from 'react';
import { AlertPriority } from '../Alerts/IAlerts';
import * as strings from 'AlertBannerApplicationCustomizerStrings';

export interface ISharePointSelectOption {
  value: string;
  label: string;
}

/**
 * Custom hook to generate priority options for dropdown/select components
 */
export const usePriorityOptions = (): ISharePointSelectOption[] => {
  return useMemo(() => ([
    {
      value: AlertPriority.Low,
      label: strings.CreateAlertPriorityLowDescription
    },
    {
      value: AlertPriority.Medium,
      label: strings.CreateAlertPriorityMediumDescription
    },
    {
      value: AlertPriority.High,
      label: strings.CreateAlertPriorityHighDescription
    },
    {
      value: AlertPriority.Critical,
      label: strings.CreateAlertPriorityCriticalDescription
    }
  ]), []);
};
