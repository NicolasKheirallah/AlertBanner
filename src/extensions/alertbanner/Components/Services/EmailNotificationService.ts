/**
 * @file EmailNotificationService.ts
 * @description Service for sending email notifications about alerts
 *   using the Microsoft Graph API `sendMail` endpoint.
 */

import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IAlertItem, AlertPriority } from "../Alerts/IAlerts";
import { logger } from "./LoggerService";

/** Configuration for the email notification service */
export interface IEmailConfig {
  /** Service account email address used as the sender */
  serviceAccountEmail: string;
  /** Graph client for making API calls */
  graphClient: MSGraphClientV3;
}

/** Email recipient */
interface IEmailRecipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

/** Graph API sendMail payload */
interface ISendMailPayload {
  message: {
    subject: string;
    body: {
      contentType: "HTML" | "Text";
      content: string;
    };
    toRecipients: IEmailRecipient[];
    importance: "low" | "normal" | "high";
  };
  saveToSentItems: boolean;
}

/**
 * Maps alert priority to email importance level.
 * @param priority - Alert priority enum value
 * @returns Graph API importance string
 */
function mapPriorityToImportance(
  priority: AlertPriority,
): "low" | "normal" | "high" {
  switch (priority) {
    case AlertPriority.Critical:
    case AlertPriority.High:
      return "high";
    case AlertPriority.Medium:
      return "normal";
    case AlertPriority.Low:
      return "low";
    default:
      return "normal";
  }
}

/**
 * Renders an alert as an HTML email body.
 * @param alert - Alert item to render
 * @returns HTML string
 */
function renderAlertEmail(alert: IAlertItem): string {
  const priorityColors: Record<string, string> = {
    critical: "#d13438",
    high: "#f7630c",
    medium: "#0078d4",
    low: "#107c10",
  };

  const borderColor = priorityColors[alert.priority] || "#0078d4";

  return `
    <div style="font-family: 'Segoe UI', sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="border-left: 4px solid ${borderColor}; padding: 16px 20px; background: #faf9f8; border-radius: 4px;">
        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
          <span style="text-transform: uppercase; font-size: 11px; font-weight: 600; color: ${borderColor}; padding: 2px 8px; border-radius: 10px; border: 1px solid ${borderColor};">${alert.priority}</span>
          <span style="font-size: 11px; color: #605e5c;">${alert.AlertType}</span>
        </div>
        <h2 style="margin: 0 0 8px 0; font-size: 18px; color: #323130;">${alert.title}</h2>
        <div style="font-size: 14px; color: #605e5c; line-height: 1.5;">${alert.description}</div>
        ${
          alert.linkUrl
            ? `<p style="margin: 12px 0 0;"><a href="${alert.linkUrl}" style="color: #0078d4; text-decoration: none;">${alert.linkDescription || alert.linkUrl}</a></p>`
            : ""
        }
      </div>
      <p style="font-size: 11px; color: #a19f9d; margin-top: 12px;">
        This is an automated notification from the Alert Banner system.
      </p>
    </div>
  `;
}

/**
 * Service that sends email notifications for alerts via Microsoft Graph API.
 *
 * Uses `POST /users/{serviceAccountEmail}/sendMail` to send notifications
 * to targeted users when high-priority alerts are created or updated.
 */
export class EmailNotificationService {
  private config: IEmailConfig;

  constructor(config: IEmailConfig) {
    this.config = config;
  }

  /**
   * Sends an email notification for an alert to the specified recipients.
   *
   * @param alert - The alert to notify about
   * @param recipientEmails - Array of email addresses to notify
   * @returns Promise that resolves when the email is sent
   * @throws Error if Graph API call fails
   */
  public async sendAlertNotification(
    alert: IAlertItem,
    recipientEmails: string[],
  ): Promise<void> {
    if (!recipientEmails.length) {
      logger.debug(
        "EmailNotificationService",
        "No recipients provided, skipping email",
      );
      return;
    }

    if (!this.config.serviceAccountEmail) {
      logger.warn(
        "EmailNotificationService",
        "No service account email configured, skipping email",
      );
      return;
    }

    const toRecipients: IEmailRecipient[] = recipientEmails.map((email) => ({
      emailAddress: { address: email },
    }));

    const payload: ISendMailPayload = {
      message: {
        subject: `[Alert - ${alert.priority.toUpperCase()}] ${alert.title}`,
        body: {
          contentType: "HTML",
          content: renderAlertEmail(alert),
        },
        toRecipients,
        importance: mapPriorityToImportance(alert.priority),
      },
      saveToSentItems: false,
    };

    try {
      await this.config.graphClient
        .api(`/users/${this.config.serviceAccountEmail}/sendMail`)
        .post(payload);

      logger.info(
        "EmailNotificationService",
        `Email sent for alert "${alert.title}" to ${recipientEmails.length} recipient(s)`,
      );
    } catch (error) {
      logger.error(
        "EmailNotificationService",
        "Failed to send email notification",
        error,
      );
      throw error;
    }
  }

  /**
   * Sends a test email to verify the service account configuration.
   *
   * @param recipientEmail - Email address to send the test to
   * @returns Promise that resolves when the test email is sent
   */
  public async sendTestEmail(recipientEmail: string): Promise<void> {
    const payload: ISendMailPayload = {
      message: {
        subject: "Alert Banner — Email Configuration Test",
        body: {
          contentType: "HTML",
          content: `
            <div style="font-family: 'Segoe UI', sans-serif; max-width: 600px; margin: 0 auto;">
              <div style="padding: 16px 20px; background: #dff6dd; border-radius: 4px; border: 1px solid #c7e6d0;">
                <h2 style="margin: 0 0 8px 0; font-size: 18px; color: #107c10;">✓ Configuration Successful</h2>
                <p style="font-size: 14px; color: #323130; margin: 0;">
                  Your email notification service is configured correctly.
                  Alerts with email notification will be sent from <strong>${this.config.serviceAccountEmail}</strong>.
                </p>
              </div>
            </div>
          `,
        },
        toRecipients: [{ emailAddress: { address: recipientEmail } }],
        importance: "normal",
      },
      saveToSentItems: false,
    };

    await this.config.graphClient
      .api(`/users/${this.config.serviceAccountEmail}/sendMail`)
      .post(payload);

    logger.info(
      "EmailNotificationService",
      `Test email sent to ${recipientEmail}`,
    );
  }
}
