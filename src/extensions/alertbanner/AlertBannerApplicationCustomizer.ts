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
import { MSGraphClientV3, SPHttpClient } from "@microsoft/sp-http";
import { AlertsProvider } from "./Components/Context/AlertsContext";
import { LocalizationService } from "./Components/Services/LocalizationService";
import { LocalizationProvider } from "./Components/Hooks/useLocalization";
import Alerts from "./Components/Alerts/Alerts";
import { logger } from './Components/Services/LoggerService';
import { SiteContextService } from "./Components/Services/SiteContextService";
import { setIconOptions } from "@fluentui/style-utilities";

export default class AlertsBannerApplicationCustomizer extends BaseApplicationCustomizer<IAlertsBannerApplicationCustomizerProperties> {
  private static readonly COMPONENT_ID = "4b274e80-896b-4c87-9a78-d751d9dff522";
  private static readonly SETTINGS_PERSIST_DEBOUNCE_MS = 800;
  private _topPlaceholderContent: PlaceholderContent | undefined;
  private _customProperties: IAlertsBannerApplicationCustomizerProperties;
  private _siteIds: string[] | null = null; // Cache site IDs to prevent recalculation
  private _isRendering: boolean = false; // Prevent concurrent renders
  private _lastRenderedSiteId: string | null = null; // Track last site to detect SPA navigation
  private _settingsPersistDebounceId: number | null = null;

  @override
  public async onInit(): Promise<void> {
    // Suppress duplicate icon registration warnings from Fluent UI v8 dependencies
    setIconOptions({ disableWarnings: true });

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

    // DISABLED BY DEFAULT - target site selection can be enabled in settings
    this._customProperties.enableTargetSite =
      this._customProperties.enableTargetSite !== undefined ?
      this._customProperties.enableTargetSite : false;

    this._customProperties.emailServiceAccount =
      this._customProperties.emailServiceAccount !== undefined ?
      this._customProperties.emailServiceAccount : "";

    this._customProperties.copilotEnabled =
      this._customProperties.copilotEnabled !== undefined ?
      this._customProperties.copilotEnabled : false;

    this._loadSettingsSnapshot();
    this._persistCustomProperties();
    this._persistSettingsSnapshot();
  }

  private _persistCustomProperties(): void {
    this.properties.alertTypesJson = this._customProperties.alertTypesJson;
    this.properties.userTargetingEnabled = this._customProperties.userTargetingEnabled;
    this.properties.notificationsEnabled = this._customProperties.notificationsEnabled;
    this.properties.enableTargetSite = this._customProperties.enableTargetSite;
    this.properties.emailServiceAccount = this._customProperties.emailServiceAccount;
    this.properties.copilotEnabled = this._customProperties.copilotEnabled;
  }

  @override
  public onDispose(): void {
    if (this._settingsPersistDebounceId) {
      window.clearTimeout(this._settingsPersistDebounceId);
      this._settingsPersistDebounceId = null;
    }
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
    enableTargetSite: boolean;
    emailServiceAccount?: string;
    copilotEnabled?: boolean;
  }): void => {
    const hasChanged =
      this._customProperties.alertTypesJson !== settings.alertTypesJson ||
      this._customProperties.userTargetingEnabled !== settings.userTargetingEnabled ||
      this._customProperties.notificationsEnabled !== settings.notificationsEnabled ||
      this._customProperties.enableTargetSite !== settings.enableTargetSite ||
      this._customProperties.emailServiceAccount !== settings.emailServiceAccount ||
      this._customProperties.copilotEnabled !== settings.copilotEnabled;

    if (!hasChanged) {
      return;
    }

    this._customProperties = {
      ...this._customProperties,
      alertTypesJson: settings.alertTypesJson,
      userTargetingEnabled: settings.userTargetingEnabled,
      notificationsEnabled: settings.notificationsEnabled,
      enableTargetSite: settings.enableTargetSite,
      emailServiceAccount: settings.emailServiceAccount,
      copilotEnabled: settings.copilotEnabled,
    };

    this._persistCustomProperties();
    this._scheduleSettingsPersistence();
  };

  private _getSettingsSnapshotKey(): string {
    return `alertbanner-settings-${this.context.pageContext.site.id.toString()}`;
  }

  private _loadSettingsSnapshot(): void {
    if (typeof window === "undefined") {
      return;
    }

    try {
      const raw = window.localStorage.getItem(this._getSettingsSnapshotKey());
      if (!raw) {
        return;
      }

      const parsed = JSON.parse(raw) as Partial<IAlertsBannerApplicationCustomizerProperties>;

      if (typeof parsed.alertTypesJson === "string") {
        this._customProperties.alertTypesJson = parsed.alertTypesJson;
      }
      if (typeof parsed.userTargetingEnabled === "boolean") {
        this._customProperties.userTargetingEnabled = parsed.userTargetingEnabled;
      }
      if (typeof parsed.notificationsEnabled === "boolean") {
        this._customProperties.notificationsEnabled = parsed.notificationsEnabled;
      }
      if (typeof parsed.enableTargetSite === "boolean") {
        this._customProperties.enableTargetSite = parsed.enableTargetSite;
      }
      if (typeof parsed.emailServiceAccount === "string") {
        this._customProperties.emailServiceAccount = parsed.emailServiceAccount;
      }
      if (typeof parsed.copilotEnabled === "boolean") {
        this._customProperties.copilotEnabled = parsed.copilotEnabled;
      }
    } catch (error) {
      logger.warn(
        "ApplicationCustomizer",
        "Failed to load settings snapshot from localStorage",
        error,
      );
    }
  }

  private _persistSettingsSnapshot(): void {
    if (typeof window === "undefined") {
      return;
    }

    try {
      window.localStorage.setItem(
        this._getSettingsSnapshotKey(),
        JSON.stringify({
          alertTypesJson: this._customProperties.alertTypesJson,
          userTargetingEnabled: this._customProperties.userTargetingEnabled,
          notificationsEnabled: this._customProperties.notificationsEnabled,
          enableTargetSite: this._customProperties.enableTargetSite,
          emailServiceAccount: this._customProperties.emailServiceAccount,
          copilotEnabled: this._customProperties.copilotEnabled,
        }),
      );
    } catch (error) {
      logger.warn(
        "ApplicationCustomizer",
        "Failed to persist settings snapshot to localStorage",
        error,
      );
    }
  }

  private _scheduleSettingsPersistence(): void {
    this._persistSettingsSnapshot();

    if (typeof window === "undefined") {
      return;
    }

    if (this._settingsPersistDebounceId) {
      window.clearTimeout(this._settingsPersistDebounceId);
    }

    this._settingsPersistDebounceId = window.setTimeout(() => {
      this._settingsPersistDebounceId = null;
      void this._persistCustomActionProperties();
    }, AlertsBannerApplicationCustomizer.SETTINGS_PERSIST_DEBOUNCE_MS);
  }

  private async _persistCustomActionProperties(): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${webUrl}/_api/web/UserCustomActions?$select=Id,ClientSideComponentId`;
      const listResponse = await this.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
          },
        },
      );

      if (!listResponse.ok) {
        logger.warn(
          "ApplicationCustomizer",
          "Failed to query user custom actions for settings persistence",
          { status: listResponse.status },
        );
        return;
      }

      const actions = (await listResponse.json()) as {
        value?: Array<{ Id: string; ClientSideComponentId?: string }>;
      };

      const customAction = actions.value?.find(
        (action) =>
          (action.ClientSideComponentId || "")
            .replace(/[{}]/g, "")
            .toLowerCase() ===
          AlertsBannerApplicationCustomizer.COMPONENT_ID,
      );

      if (!customAction?.Id) {
        logger.warn(
          "ApplicationCustomizer",
          "No matching custom action found for settings persistence",
        );
        return;
      }

      const updateUrl = `${webUrl}/_api/web/UserCustomActions(guid'${customAction.Id}')`;
      const updateResponse = await this.context.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({
            __metadata: { type: "SP.UserCustomAction" },
            ClientSideComponentProperties: JSON.stringify({
              alertTypesJson: this._customProperties.alertTypesJson,
              userTargetingEnabled: this._customProperties.userTargetingEnabled,
              notificationsEnabled: this._customProperties.notificationsEnabled,
              enableTargetSite: this._customProperties.enableTargetSite,
              emailServiceAccount: this._customProperties.emailServiceAccount || "",
              copilotEnabled: !!this._customProperties.copilotEnabled,
            }),
          }),
        },
      );

      if (!updateResponse.ok) {
        logger.warn(
          "ApplicationCustomizer",
          "Failed to persist custom action settings",
          { status: updateResponse.status },
        );
      }
    } catch (error) {
      logger.warn(
        "ApplicationCustomizer",
        "Error while persisting custom action settings",
        error,
      );
    }
  }

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
            enableTargetSite: this._customProperties.enableTargetSite,
            emailServiceAccount: this._customProperties.emailServiceAccount,
            copilotEnabled: this._customProperties.copilotEnabled,
            onSettingsChange: this._handleSettingsChange
          }
        );

        // Wrap with the LocalizationProvider and AlertsProvider
        const alertsApp = React.createElement(
          LocalizationProvider,
          { children: React.createElement(AlertsProvider, { children: alertsComponent }) }
        );

        // Render with error handling
        // NOTE: Using ReactDOM.render is required for SPFx compatibility (React 17)
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
