import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IAlertProps {
  showRemoteAlerts?: boolean;
  remoteAlertsSource?: string;
  siteIds?: string[]; // Optional property
  graphClient: MSGraphClientV3;  // Pass the initialized MSGraphClientV3
}

export interface IAlertState {
  alerts: IAlertItem[];
}

export interface IAlertItem {
  Id: number;
  title: string;
  description: string;
  AlertType: AlertType; // Typed as AlertType enum
  link?: {
    Url: string;
    Description: string;
  };
}

export enum AlertType {
  Info = "Info",
  Warning = "Warning",
  Maintenance = "Maintenance",
  Interruption = "Interruption",
}
