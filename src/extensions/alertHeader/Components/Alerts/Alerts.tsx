// Alerts.tsx

import * as React from 'react';
import styles from './Alerts.module.scss';
import { IAlertProps, IAlertState, IAlertItem, AlertType } from './IAlerts.types';
import AlertItem from '../AlertItem/AlertItem';

class Alerts extends React.Component<IAlertProps, IAlertState> {
  public static readonly LIST_TITLE = "Alerts";

  public state: IAlertState = {
    alerts: [],
  };

  public async componentDidMount(): Promise<void> {
    try {
      const allAlerts: IAlertItem[] = [];

      // Fetch alerts from site IDs if provided
      if (this.props.siteIds && this.props.siteIds.length > 0) {
        for (const siteId of this.props.siteIds) {
          const siteAlerts = await this.fetchAlerts(siteId);
          allAlerts.push(...siteAlerts);
        }
      }

      // Fetch remote alerts if specified
      if (this.props.showRemoteAlerts && this.props.remoteAlertsSource) {
        const remoteAlerts = await this.fetchRemoteAlerts(this.props.remoteAlertsSource);
        allAlerts.push(...remoteAlerts);
      }

      // If no alerts were fetched, handle accordingly
      if (allAlerts.length === 0) {
        console.warn('No alerts fetched from any source.');
        // Optionally, you can set an empty alerts array or handle as needed
      }

      // Proceed with the rest of your logic (caching, filtering, updating state)
      const uniqueAlerts = this._removeDuplicateAlerts(allAlerts);

      const cachedAlerts = this._getFromLocalStorage('AllAlerts');

      const alertsAreDifferent = this._areAlertsDifferent(uniqueAlerts, cachedAlerts);

      if (alertsAreDifferent) {
        this._saveToLocalStorage('AllAlerts', uniqueAlerts);
      }

      const alertsToShow = alertsAreDifferent ? uniqueAlerts : cachedAlerts || [];
      const closedAlerts = this._getClosedAlerts();
      const filteredAlerts = alertsToShow.filter(alert => !closedAlerts.includes(alert.Id));

      this.setState({ alerts: filteredAlerts });
    } catch (error) {
      console.error('Error initializing alerts:', error);
    }
  }

  private async fetchAlerts(siteId: string): Promise<IAlertItem[]> {
    const dateTimeNow = new Date().toISOString();
    const filterQuery = `fields/StartDateTime le '${dateTimeNow}' and fields/EndDateTime ge '${dateTimeNow}'`;

    try {
      const response = await this.props.graphClient
        .api(`/sites/${siteId}/lists/${Alerts.LIST_TITLE}/items`)
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .expand('fields($select=Title,AlertType,Description,Link,StartDateTime,EndDateTime)')
        .filter(filterQuery)
        .orderby('fields/StartDateTime desc')
        .get();

      return response.value.map((item: any) => ({
        Id: parseInt(item.id, 10),
        title: item.fields.Title,
        description: item.fields.Description,
        AlertType: item.fields.AlertType as AlertType,
        link: item.fields.Link,
      }));
    } catch (error) {
      console.error(`Error fetching alerts from site ${siteId}:`, error);
      return [];
    }
  }

  private async fetchRemoteAlerts(sourceUrl: string): Promise<IAlertItem[]> {
    try {
      const response = await fetch(sourceUrl);
      const data = await response.json();
      // Transform data into IAlertItem[]
      const alerts: IAlertItem[] = data.map((item: any) => ({
        Id: item.Id,
        title: item.Title,
        description: item.Description,
        AlertType: item.AlertType as AlertType,
        link: item.Link,
      }));
      return alerts;
    } catch (error) {
      console.error('Error fetching remote alerts:', error);
      return [];
    }
  }

  private _removeDuplicateAlerts(alerts: IAlertItem[]): IAlertItem[] {
    const seenIds = new Set<number>();
    return alerts.filter(alert => {
      if (seenIds.has(alert.Id)) {
        return false;
      } else {
        seenIds.add(alert.Id);
        return true;
      }
    });
  }

  private _areAlertsDifferent(newAlerts: IAlertItem[], cachedAlerts: IAlertItem[] | null): boolean {
    if (!cachedAlerts) return true;
    if (newAlerts.length !== cachedAlerts.length) return true;

    for (let i = 0; i < newAlerts.length; i++) {
      const newAlert = newAlerts[i];
      const cachedAlert = cachedAlerts[i];

      if (
        newAlert.Id !== cachedAlert.Id ||
        newAlert.title !== cachedAlert.title ||
        newAlert.description !== cachedAlert.description ||
        newAlert.AlertType !== cachedAlert.AlertType ||
        newAlert.link?.Url !== cachedAlert.link?.Url
      ) {
        return true;
      }
    }

    return false;
  }

  private _removeAlert = (id: number): void => {
    this.setState((prevState) => {
      const updatedAlerts = prevState.alerts.filter((alert) => alert.Id !== id);
      this._addClosedAlert(id);
      return { alerts: updatedAlerts };
    });
  };

  private _getClosedAlerts(): number[] {
    const stored = this._getFromSessionStorage('ClosedAlerts');
    return Array.isArray(stored) ? stored : []; // Ensure the result is always an array
  }

  private _addClosedAlert(id: number): void {
    const closedAlerts = this._getClosedAlerts();
    if (!closedAlerts.includes(id)) {
      closedAlerts.push(id);
      this._saveToSessionStorage('ClosedAlerts', closedAlerts);
    }
  }

  private _getFromLocalStorage(key: string): IAlertItem[] | null {
    try {
      const data = localStorage.getItem(key);
      return data ? JSON.parse(data) : null;
    } catch (error) {
      console.error('Error accessing localStorage:', error);
      return null;
    }
  }

  private _saveToLocalStorage(key: string, data: IAlertItem[]): void {
    try {
      localStorage.setItem(key, JSON.stringify(data));
    } catch (error) {
      console.error('Error saving to localStorage:', error);
    }
  }

  private _getFromSessionStorage(key: string): number[] {
    try {
      const data = sessionStorage.getItem(key);
      return data ? JSON.parse(data) : [];
    } catch (error) {
      console.error('Error accessing sessionStorage:', error);
      return [];
    }
  }

  private _saveToSessionStorage(key: string, data: number[]): void {
    try {
      sessionStorage.setItem(key, JSON.stringify(data));
    } catch (error) {
      console.error('Error saving to sessionStorage:', error);
    }
  }

  public render(): React.ReactElement<IAlertProps> {
    return (
      <div className={styles.alerts}>
        <div className={styles.container}>
          {this.state.alerts.map((alert) => (
            <AlertItem
              key={alert.Id}
              item={alert}
              remove={this._removeAlert}
            />
          ))}
        </div>
      </div>
    );
  }
}

export default Alerts;
