
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IAlertProps {
  showRemoteAlerts?: boolean;
  remoteAlertsSource?: string;
  siteId?: string;
  graphClient: MSGraphClientV3;  // Pass the initialized MSGraphClientV3
}

export interface IAlertState {
  alerts: Array<IAlertItem>;
}

export interface IAlertItem {
  Id: number;
  title: string;
  description: string;
  AlertType: string;
  link: IAlertLink;
}

export interface IAlertLink {
  Url: string;
  Description: string;
}

export enum AlertType {
  Info = "Info",
  Warning = "Warning",
  Maintenance = "Maintenance",
  Interruption = "Interruption",
}
