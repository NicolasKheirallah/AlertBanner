import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import Alerts from "./Components/Alerts/Alerts";
import { IAlertProps } from "./Components/Alerts/IAlerts.types";
import { MSGraphClientV3 } from "@microsoft/sp-http";

// Interface for any properties passed to the application customizer (currently empty)
export interface IAlertsHeaderApplicationCustomizerProperties {}

// Main class for the Alerts Header Application Customizer
export default class AlertsHeaderApplicationCustomizer extends BaseApplicationCustomizer<IAlertsHeaderApplicationCustomizerProperties> {
  // Placeholder for the top area of the SharePoint page
  private _topPlaceholderContent: PlaceholderContent | undefined;

  // onInit is a lifecycle method that runs when the application customizer is initialized
  @override
  public async onInit(): Promise<void> {
    // Register an event listener to re-render the placeholders if they change
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderTopPlaceholder
    );

    // Initial call to render placeholders
    await this._renderTopPlaceholder();
  }

  // onDispose is a lifecycle method that runs when the application customizer is disposed
  @override
  public onDispose(): void {
    // Remove the event listener when the customizer is disposed
    this.context.placeholderProvider.changedEvent.remove(
      this,
      this._renderTopPlaceholder
    );
    this._disposeAlertsComponent();
    super.onDispose();
  }

  // Method to render the top placeholder; called when placeholders change or during initialization
  private async _renderTopPlaceholder(): Promise<void> {
    if (!this._topPlaceholderContent) {
      // Check if the Top placeholder is available
      if (
        !this.context.placeholderProvider.placeholderNames.includes(
          PlaceholderName.Top
        )
      ) {
        console.warn("Top placeholder is not available.");
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

  // Private method to render the Alerts component in the top placeholder
  private async _renderAlertsComponent(): Promise<void> {
    try {
      if (
        this._topPlaceholderContent &&
        this._topPlaceholderContent.domElement
      ) {
        // Get the Microsoft Graph client (version 3) for making API calls
        const msGraphClient: MSGraphClientV3 = (await this.context.msGraphClientFactory.getClient(
          "3"
        )) as MSGraphClientV3;

        // Get the current site ID
        const currentSiteId: string = this.context.pageContext.site.id.toString();

        // Get the hub site ID, if available
        const hubSiteId: string = this.context.pageContext.legacyPageContext.hubSiteId.toString();

        // Get the SharePoint home site ID
        const homeSiteResponse = await msGraphClient
          .api("/sites/root")
          .select("id")
          .get();
        const homeSiteId: string = homeSiteResponse.id;

        // Prepare the array of site IDs, ensuring uniqueness
        const siteIds: string[] = [currentSiteId];

        if (
          hubSiteId &&
          hubSiteId !== "00000000-0000-0000-0000-000000000000" &&
          hubSiteId !== currentSiteId &&
          !siteIds.includes(hubSiteId)
        ) {
          siteIds.push(hubSiteId);
        }

        if (
          homeSiteId &&
          homeSiteId !== currentSiteId &&
          homeSiteId !== hubSiteId &&
          !siteIds.includes(homeSiteId)
        ) {
          siteIds.push(homeSiteId);
        }

        // Create the Alerts React element with the necessary props
        const alertsComponentElement: React.ReactElement<IAlertProps> = React.createElement(
          Alerts,
          {
            siteIds: siteIds, // Pass the array of site IDs
            graphClient: msGraphClient, // Pass the Graph client to the Alerts component
          }
        );

        // Render the Alerts component into the top placeholder's DOM element
        ReactDOM.render(
          alertsComponentElement,
          this._topPlaceholderContent.domElement
        );
      }
    } catch (error) {
      console.error("Error rendering Alerts component:", error);
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
