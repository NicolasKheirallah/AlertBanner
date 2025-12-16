import * as React from "react";
import * as ReactDOM from "react-dom";
// Version 4.2.0 - Fixed AttachmentFiles Graph API errors and language filtering
import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { IAlertsBannerApplicationCustomizerProperties } from "./Components/Alerts/IAlerts";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { AlertsProvider } from "./Components/Context/AlertsContext";
import { LocalizationService } from "./Components/Services/LocalizationService";
import { LocalizationProvider } from "./Components/Hooks/useLocalization";
import Alerts from "./Components/Alerts/Alerts";
import { logger } from './Components/Services/LoggerService';
import StorageService from './Components/Services/StorageService';
import { SiteContextService } from "./Components/Services/SiteContextService";

export default class AlertsBannerApplicationCustomizer extends BaseApplicationCustomizer<IAlertsBannerApplicationCustomizerProperties> {
  private _topPlaceholderContent: PlaceholderContent | undefined;
  private _customProperties: IAlertsBannerApplicationCustomizerProperties;
  private _siteIds: string[] | null = null; // Cache site IDs to prevent recalculation
  private _isRendering: boolean = false; // Prevent concurrent renders
  private _lastRenderedSiteId: string | null = null; // Track last site to detect SPA navigation
  private readonly _storageService: StorageService = StorageService.getInstance();

  @override
  public async onInit(): Promise<void> {
    // Initialize localization service
    const localizationService = LocalizationService.getInstance(this.context);
    await localizationService.initialize(this.context);

    // Initialize default configuration
    this._initializeDefaultProperties();

    // Add listener for placeholder changes
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderTopPlaceholder
    );

    await this._renderTopPlaceholder();
  }

  private _initializeDefaultProperties(): void {
    // Instead of modifying this.properties directly, create a local copy
    this._customProperties = { ...this.properties };

    // Merge persisted settings from storage if available
    const persistedSettings = this._storageService.getFromLocalStorage<IAlertsBannerApplicationCustomizerProperties>('AlertBannerSettings');
    if (persistedSettings) {
      this._customProperties = {
        ...this._customProperties,
        ...persistedSettings
      };
    }

    // Set default alert types if none are provided
    if (!this._customProperties.alertTypesJson || this._customProperties.alertTypesJson === "[]") {
      const defaultAlertTypes = [
        {
          "name": "Info",
          "iconName": "Info",
          "backgroundColor": "#389899",
          "textColor": "#ffffff",
          "additionalStyles": "",
          "priorityStyles": {
            "critical": "border: 2px solid #E81123;",
            "high": "border: 1px solid #EA4300;",
            "medium": "",
            "low": ""
          }
        },
        {
          "name": "Warning",
          "iconName": "Warning",
          "backgroundColor": "#f1c40f",
          "textColor": "#000000",
          "additionalStyles": "",
          "priorityStyles": {
            "critical": "border: 2px solid #E81123;",
            "high": "border: 1px solid #EA4300;",
            "medium": "",
            "low": ""
          }
        },
        {
          "name": "Maintenance",
          "iconName": "ConstructionCone",
          "backgroundColor": "#afd6d6",
          "textColor": "#000000",
          "additionalStyles": "",
          "priorityStyles": {
            "critical": "border: 2px solid #E81123;",
            "high": "border: 1px solid #EA4300;",
            "medium": "",
            "low": ""
          }
        },
        {
          "name": "Interruption",
          "iconName": "Error",
          "backgroundColor": "#c54644",
          "textColor": "#ffffff",
          "additionalStyles": "",
          "priorityStyles": {
            "critical": "border: 2px solid #E81123;",
            "high": "border: 1px solid #EA4300;",
            "medium": "",
            "low": ""
          }
        }
      ];

      this._customProperties.alertTypesJson = JSON.stringify(defaultAlertTypes);
    }

    // Set defaults for any missing properties
    this._customProperties.userTargetingEnabled =
      this._customProperties.userTargetingEnabled !== undefined ?
      this._customProperties.userTargetingEnabled : true;

    // DISABLED BY DEFAULT - notifications can be enabled in settings
    this._customProperties.notificationsEnabled =
      this._customProperties.notificationsEnabled !== undefined ?
      this._customProperties.notificationsEnabled : false;

    this._persistCustomProperties();
  }

  private _persistCustomProperties(): void {
    this.properties.alertTypesJson = this._customProperties.alertTypesJson;
    this.properties.userTargetingEnabled = this._customProperties.userTargetingEnabled;
    this.properties.notificationsEnabled = this._customProperties.notificationsEnabled;

    this._storageService.saveToLocalStorage('AlertBannerSettings', this._customProperties);
  }

  @override
  public onDispose(): void {
    this.context.placeholderProvider.changedEvent.remove(
      this,
      this._renderTopPlaceholder
    );
    this._disposeAlertsComponent();
    super.onDispose();
  }

  private async _renderTopPlaceholder(): Promise<void> {
    if (!this._topPlaceholderContent) {
      if (
        !this.context.placeholderProvider.placeholderNames.includes(
          PlaceholderName.Top
        )
      ) {
        logger.warn('ApplicationCustomizer', 'Top placeholder is not available');
        return;
      }

      this._topPlaceholderContent = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._disposeAlertsComponent }
      );
    }

    if (this._topPlaceholderContent) {
      await this._renderAlertsComponent();
    }
  }

  private _handleSettingsChange = (settings: {
    alertTypesJson: string;
    userTargetingEnabled: boolean;
    notificationsEnabled: boolean;
  }): void => {
    const hasChanged =
      this._customProperties.alertTypesJson !== settings.alertTypesJson ||
      this._customProperties.userTargetingEnabled !== settings.userTargetingEnabled ||
      this._customProperties.notificationsEnabled !== settings.notificationsEnabled;

    if (!hasChanged) {
      return;
    }

    this._customProperties = {
      ...this._customProperties,
      alertTypesJson: settings.alertTypesJson,
      userTargetingEnabled: settings.userTargetingEnabled,
      notificationsEnabled: settings.notificationsEnabled
    };

    this._persistCustomProperties();
  };

  private async _renderAlertsComponent(): Promise<void> {
    if (this._isRendering) {
      return;
    }

    try {
      this._isRendering = true;

      if (
        this._topPlaceholderContent &&
        this._topPlaceholderContent.domElement
      ) {
        // Try to get Graph client with version 3, with error handling
        let msGraphClient: MSGraphClientV3;
        try {
          msGraphClient = await this.context.msGraphClientFactory.getClient("3") as MSGraphClientV3;
        } catch (graphError) {
          logger.error('ApplicationCustomizer', 'Error getting Graph client v3', graphError);
          throw graphError; // Re-throw to be caught by outer try/catch
        }

        // Initialize SiteContextService
        const siteContextService = SiteContextService.getInstance(this.context, msGraphClient);
        await siteContextService.initialize();

        const currentSiteId = this.context.pageContext.site.id.toString();

        if (this._lastRenderedSiteId && this._lastRenderedSiteId !== currentSiteId) {
          this._siteIds = null;
        }

        this._lastRenderedSiteId = currentSiteId;

        if (!this._siteIds) {
          // Use the robust site detection from SiteContextService
          this._siteIds = siteContextService.getAlertSourceSites();
          
          logger.info('ApplicationCustomizer', 'Resolved alert source sites', { 
            sites: this._siteIds,
            homeSite: siteContextService.getHomeSite()?.url,
            hubSite: siteContextService.getHubSite()?.url,
            currentSite: siteContextService.getCurrentSite()?.url
          });
        }

        // Get alert types from our custom properties
        const alertTypesJsonString = this._customProperties.alertTypesJson || "[]";

        // Create the AlertsContext provider
        const alertsComponent = React.createElement(
          Alerts,
          {
            siteIds: this._siteIds, // Use cached site IDs
            graphClient: msGraphClient,
            context: this.context,
            alertTypesJson: alertTypesJsonString,
            userTargetingEnabled: this._customProperties.userTargetingEnabled,
            notificationsEnabled: this._customProperties.notificationsEnabled,
            onSettingsChange: this._handleSettingsChange
          }
        );

        // Wrap with the LocalizationProvider and AlertsProvider
        const alertsApp = React.createElement(
          LocalizationProvider,
          { children: React.createElement(AlertsProvider, { children: alertsComponent }) }
        );

        // Render with error handling
        ReactDOM.render(
          alertsApp,
          this._topPlaceholderContent.domElement
        );
      }
    } catch (error) {
      logger.error('ApplicationCustomizer', 'Error rendering Alerts component', error);

      // Render a minimal error message instead of failing completely
      if (this._topPlaceholderContent && this._topPlaceholderContent.domElement) {
        const errorElement = React.createElement(
          'div',
          { style: { padding: '10px', color: '#666', fontSize: '13px' } },
          'Unable to load alerts at this time. Please try refreshing the page.'
        );

        ReactDOM.render(
          errorElement,
          this._topPlaceholderContent.domElement
        );
      }
    } finally {
      this._isRendering = false;
    }
  }

  // Dispose the React component when the customizer is disposed
  private _disposeAlertsComponent = (): void => {
    if (
      this._topPlaceholderContent &&
      this._topPlaceholderContent.domElement
    ) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholderContent.domElement);
    }
  };
}
