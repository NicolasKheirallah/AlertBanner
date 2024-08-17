import * as React from 'react';
import styles from './Alerts.module.scss';
import { IAlertProps, IAlertState } from './index';
import { IAlertItem } from './IAlerts.types';
import AlertItem from '../AlertItem/AlertItem';
import { MSGraphClientV3 } from '@microsoft/sp-http';

class Alerts extends React.Component<IAlertProps, IAlertState> {
  private _graphClient: MSGraphClientV3 = this.props.graphClient;
  private _storageKey = this.props.showRemoteAlerts ? "SPFXClosedAlerts" : `${this.props.siteId}ClosedAlerts`;
  private _cacheKey = this.props.showRemoteAlerts ? "SPFXGlobalAlerts" : `${this.props.siteId}AllAlerts`;

  public static readonly LIST_TITLE = "Alerts";

  public state: IAlertState = {
    alerts: [],
  };

  public async componentDidMount(): Promise<void> {
    try {
      const cachedAlerts = this._getFromLocalStorage(this._cacheKey);
      const alerts = cachedAlerts ?? await this.fetchAlerts();

      if (!cachedAlerts) {
        this._saveToLocalStorage(this._cacheKey, alerts);
      }

      const closedAlerts = this._getClosedAlerts();
      const filteredAlerts = alerts.filter((alert) => !closedAlerts.includes(alert.Id));

      this.setState({ alerts: filteredAlerts });
    } catch (error) {
      console.error('Error initializing alerts:', error);
    }
  }

  private async fetchAlerts(): Promise<IAlertItem[]> {
    const dateTimeNow = new Date().toISOString();
    const filterQuery = `fields/StartDateTime le '${dateTimeNow}' and fields/EndDateTime ge '${dateTimeNow}'`;

    try {
      const response = await this._graphClient
        .api(`/sites/${this.props.siteId}/lists/${Alerts.LIST_TITLE}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand('fields($select=Title,AlertType,Description,Link,StartDateTime,EndDateTime)')
        .filter(filterQuery)
        .orderby('fields/StartDateTime desc')
        .get();

      return response.value.map((item: any) => ({
        Id: item.id,
        title: item.fields.Title,
        description: item.fields.Description,
        AlertType: item.fields.AlertType,
        link: item.fields.Link,
      }));
    } catch (error) {
      console.error('Error fetching alerts:', error);
      return [];
    }
  }

  private _removeAlert = (id: number): void => {
    this.setState((prevState) => {
      const updatedAlerts = prevState.alerts.filter((alert) => alert.Id !== id);
      this._addClosedAlert(id);
      return { alerts: updatedAlerts };
    });
  };

  private _getClosedAlerts(): number[] {
    const stored = this._getFromSessionStorage(this._storageKey);
    return Array.isArray(stored) ? stored : []; // Ensure the result is always an array
  }

  private _addClosedAlert(id: number): void {
    const closedAlerts = this._getClosedAlerts();
    if (!closedAlerts.includes(id)) {
      closedAlerts.push(id);
      this._saveToSessionStorage(this._storageKey, closedAlerts);
    }
  }

  private _getFromLocalStorage(key: string): IAlertItem[] | null {
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : null;
  }

  private _saveToLocalStorage(key: string, data: IAlertItem[]): void {
    localStorage.setItem(key, JSON.stringify(data));
  }

  private _getFromSessionStorage(key: string): number[] {
    const data = sessionStorage.getItem(key);
    return data ? JSON.parse(data) : [];
  }

  private _saveToSessionStorage(key: string, data: number[]): void {
    sessionStorage.setItem(key, JSON.stringify(data));
  }

  public render(): React.ReactElement<IAlertProps> {
    return (
      <div className={styles.alerts}>
        <div className={styles.container}>
          {this.state.alerts.map((alert) => (
            <AlertItem key={alert.Id} item={alert} remove={this._removeAlert} />
          ))}
        </div>
      </div>
    );
  }
}

export default React.memo(Alerts);
