import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import Alerts from "./Components/Alerts/Alerts";
import { IAlertProps } from "./Components/Alerts";

// Interface for any properties passed to the application customizer (currently empty)
export interface IAlertsHeaderApplicationCustomizerProperties {}

// Main class for the Alerts Header Application Customizer
export default class AlertsHeaderApplicationCustomizer extends BaseApplicationCustomizer<IAlertsHeaderApplicationCustomizerProperties> {
  // Placeholder for the top area of the SharePoint page
  private topPlaceholder: PlaceholderContent | undefined;

  // onInit is a lifecycle method that runs when the application customizer is initialized
  @override
  public async onInit(): Promise<void> {
    // Register an event listener to re-render the placeholders if they change
    this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceholders);

    // Initial call to render placeholders
    await this.renderPlaceholders();

    return Promise.resolve();
  }

  // onDispose is a lifecycle method that runs when the application customizer is disposed
  @override
  public onDispose(): void {
    // Remove the event listener when the customizer is disposed
    this.context.placeholderProvider.changedEvent.remove(this, this.renderPlaceholders);
    this.disposeControls();
    super.onDispose();
  }

  // Method to render placeholders; called when placeholders change or during initialization
  private async renderPlaceholders(): Promise<void> {
    if (!this.topPlaceholder) {
      this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
      if (this.topPlaceholder) {
        await this.renderControls(); // Ensure the Promise is handled
      }
    }
  }

  // Private method to render the Alerts component in the top placeholder
  private async renderControls(): Promise<void> {
    // Ensure that the controls are rendered only once and that the top placeholder exists
    if (this.topPlaceholder && this.topPlaceholder.domElement) {
      // Get the Microsoft Graph client (version 3) for making API calls
      const graphClient = await this.context.msGraphClientFactory.getClient('3');

      // Create the Alerts React element with the necessary props
      const alertElement: React.ReactElement<IAlertProps> = React.createElement(
        Alerts,
        {
          siteId: this.context.pageContext.site.id.toString(), // Pass the current site ID as a prop
          showRemoteAlerts: true,  // Indicate that remote alerts should be shown
          graphClient: graphClient,  // Pass the Graph client to the Alerts component
        }
      );

      // Render the Alerts component into the top placeholder's DOM element
      ReactDOM.render(alertElement, this.topPlaceholder.domElement);
    }
  }

  // Dispose the React component when the customizer is disposed
  private disposeControls(): void {
    if (this.topPlaceholder && this.topPlaceholder.domElement) {
      ReactDOM.unmountComponentAtNode(this.topPlaceholder.domElement);
    }
  }
}
