import { useMemo } from 'react';
import { AlertPriority } from '../Alerts/IAlerts';
import { ISharePointSelectOption } from '../UI/SharePointControls';
import * as strings from 'AlertBannerApplicationCustomizerStrings';

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
