import { NotificationType, AlertPriority } from "../Alerts/IAlerts";
import { IAlertItem } from "./SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export class NotificationService {
  private static instance: NotificationService;
  private graphClient: MSGraphClientV3;
  private hasNotificationPermission: boolean = false;

  private constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
    this.checkNotificationPermission();
  }

  public static getInstance(graphClient: MSGraphClientV3): NotificationService {
    if (!NotificationService.instance) {
      NotificationService.instance = new NotificationService(graphClient);
    }
    return NotificationService.instance;
  }

  private async checkNotificationPermission(): Promise<void> {
    if (!("Notification" in window)) {
      console.warn("This browser does not support notifications");
      return;
    }

    if (Notification.permission === "granted") {
      this.hasNotificationPermission = true;
    } else if (Notification.permission !== "denied") {
      const permission = await Notification.requestPermission();
      this.hasNotificationPermission = permission === "granted";
    }
  }

  public async sendNotification(alert: IAlertItem): Promise<void> {
    if (!alert.notificationType || alert.notificationType === NotificationType.None) {
      return;
    }

    const promises: Promise<void>[] = [];

    // Send browser notification if enabled
    if ((alert.notificationType === NotificationType.Browser || 
         alert.notificationType === NotificationType.Both) && 
        this.hasNotificationPermission) {
      promises.push(this.sendBrowserNotification(alert));
    }

    // Send email notification if enabled
    if (alert.notificationType === NotificationType.Email || 
        alert.notificationType === NotificationType.Both) {
      promises.push(this.sendEmailNotification(alert));
    }

    await Promise.all(promises);
  }

  private async sendBrowserNotification(alert: IAlertItem): Promise<void> {
    try {
      if (!this.hasNotificationPermission) {
        await this.checkNotificationPermission();
        if (!this.hasNotificationPermission) return;
      }

      // Use Fluent UI design system icons and styling
      const priorityConfig = {
        [AlertPriority.Low]: { icon: "ℹ️", color: "#107c10" },
        [AlertPriority.Medium]: { icon: "📢", color: "#0078d4" },
        [AlertPriority.High]: { icon: "⚠️", color: "#f7630c" },
        [AlertPriority.Critical]: { icon: "🚨", color: "#d13438" }
      };

      const config = priorityConfig[alert.priority] || priorityConfig[AlertPriority.Medium];
      
      const notification = new Notification(`${config.icon} ${alert.title}`, {
        body: this.stripHtml(alert.description),
        icon: this.generateFluentUIIcon(config.color, config.icon),
        tag: `alert-${alert.id}`, // Prevents duplicate notifications
        requireInteraction: alert.priority === AlertPriority.Critical, // Critical alerts require user interaction
        badge: this.generateFluentUIIcon(config.color, config.icon),
        silent: alert.priority === AlertPriority.Low // Low priority alerts are silent
      });

      // Create modern notification with Fluent UI styling
      this.createModernNotificationToast(alert, config);

      notification.onclick = () => {
        window.focus();
        if (alert.linkUrl) {
          window.open(alert.linkUrl, "_blank");
        }
        notification.close();
      };

      // Auto-dismiss based on priority
      const dismissTime = this.getNotificationDismissTime(alert.priority);
      if (dismissTime > 0) {
        setTimeout(() => notification.close(), dismissTime);
      }
    } catch (error) {
      console.error("Error sending browser notification:", error);
      // Fallback to modern toast notification
      this.createModernNotificationToast(alert, { icon: "📢", color: "#0078d4" });
    }
  }

  private async sendEmailNotification(alert: IAlertItem): Promise<void> {
    try {
      const priorityConfig = {
        [AlertPriority.Low]: { color: "#107c10", badge: "Low Priority" },
        [AlertPriority.Medium]: { color: "#0078d4", badge: "Medium Priority" },
        [AlertPriority.High]: { color: "#f7630c", badge: "High Priority" },
        [AlertPriority.Critical]: { color: "#d13438", badge: "Critical Priority" }
      };

      const config = priorityConfig[alert.priority] || priorityConfig[AlertPriority.Medium];
      
      // Modern Fluent UI email template
      const emailContent = `
        <div style="font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif; max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);">
          <div style="background: ${config.color}; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
            <h1 style="margin: 0; font-size: 20px; font-weight: 600;">${alert.title}</h1>
            <div style="margin-top: 8px; opacity: 0.9;">
              <span style="background: rgba(255, 255, 255, 0.2); padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px;">
                ${config.badge}
              </span>
            </div>
          </div>
          
          <div style="padding: 24px;">
            <div style="color: #323130; line-height: 1.6; margin-bottom: 20px;">
              ${alert.description}
            </div>
            
            ${alert.linkUrl ? `
              <div style="margin: 24px 0;">
                <a href="${alert.linkUrl}" 
                   style="display: inline-block; background: ${config.color}; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: 500; transition: all 0.2s ease;">
                  ${alert.linkDescription || 'Learn More'}
                </a>
              </div>
            ` : ''}
            
            <div style="border-top: 1px solid #edebe9; margin-top: 24px; padding-top: 16px;">
              <div style="display: flex; align-items: center; gap: 12px; color: #605e5c; font-size: 12px;">
                <div style="flex: 1;">
                  <strong>SharePoint Alert Banner</strong><br>
                  Created: ${new Date(alert.createdDate).toLocaleDateString()} by ${alert.createdBy}
                </div>
                <div style="width: 32px; height: 32px; background: ${config.color}; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 16px;">
                  📢
                </div>
              </div>
            </div>
          </div>
        </div>
      `;

      await this.graphClient.api('/me/sendMail').post({
        message: {
          subject: `${this.getPriorityEmoji(alert.priority)} Alert: ${alert.title}`,
          body: {
            contentType: "HTML",
            content: emailContent
          },
          toRecipients: [
            {
              emailAddress: {
                address: "me" // Sends to the current user
              }
            }
          ],
          importance: alert.priority === AlertPriority.Critical ? "high" : 
                     alert.priority === AlertPriority.High ? "high" : "normal"
        }
      });
    } catch (error) {
      console.error("Error sending email notification:", error);
    }
  }

  // Helper method to strip HTML for notification body
  private stripHtml(html: string): string {
    // Create a temporary element to parse HTML without rendering
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = html;
    return tempDiv.textContent || tempDiv.innerText || "";
  }

  // Generate Fluent UI style icon data URL
  private generateFluentUIIcon(color: string, emoji: string): string {
    const svg = `
      <svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
        <circle cx="16" cy="16" r="16" fill="${color}"/>
        <text x="16" y="22" text-anchor="middle" font-size="16" fill="white">${emoji}</text>
      </svg>
    `;
    return `data:image/svg+xml;base64,${btoa(svg)}`;
  }

  // Create modern toast notification with Fluent UI styling
  private createModernNotificationToast(alert: IAlertItem, config: { icon: string, color: string }): void {
    // Remove existing toast if present
    const existingToast = document.getElementById(`alert-toast-${alert.id}`);
    if (existingToast) {
      existingToast.remove();
    }

    const toast = document.createElement('div');
    toast.id = `alert-toast-${alert.id}`;
    toast.innerHTML = `
      <div style="
        position: fixed;
        top: 20px;
        right: 20px;
        max-width: 400px;
        background: white;
        border-radius: 8px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
        border-left: 4px solid ${config.color};
        z-index: 10000;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
        animation: slideInRight 0.3s cubic-bezier(0.16, 1, 0.3, 1);
      ">
        <div style="padding: 16px 20px;">
          <div style="display: flex; align-items: flex-start; gap: 12px;">
            <div style="
              width: 24px;
              height: 24px;
              background: ${config.color};
              border-radius: 50%;
              display: flex;
              align-items: center;
              justify-content: center;
              color: white;
              font-size: 14px;
              flex-shrink: 0;
            ">
              ${config.icon}
            </div>
            <div style="flex: 1; min-width: 0;">
              <div style="
                font-weight: 600;
                font-size: 14px;
                color: #323130;
                line-height: 1.3;
                margin-bottom: 4px;
              ">${alert.title}</div>
              <div style="
                font-size: 13px;
                color: #605e5c;
                line-height: 1.4;
                max-height: 60px;
                overflow: hidden;
              ">${this.stripHtml(alert.description)}</div>
              ${alert.linkUrl ? `
                <div style="margin-top: 8px;">
                  <a href="${alert.linkUrl}" 
                     target="_blank"
                     style="
                       color: ${config.color};
                       font-size: 12px;
                       text-decoration: none;
                       font-weight: 500;
                     "
                     onclick="document.getElementById('alert-toast-${alert.id}').remove();">
                    ${alert.linkDescription || 'Learn More'} →
                  </a>
                </div>
              ` : ''}
            </div>
            <button 
              onclick="document.getElementById('alert-toast-${alert.id}').remove();"
              style="
                background: none;
                border: none;
                color: #605e5c;
                cursor: pointer;
                font-size: 16px;
                padding: 0;
                width: 20px;
                height: 20px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 2px;
              "
              onmouseover="this.style.backgroundColor='#f3f2f1';"
              onmouseout="this.style.backgroundColor='transparent';"
            >×</button>
          </div>
        </div>
      </div>
    `;

    document.body.appendChild(toast);

    // Auto-dismiss after specified time
    const dismissTime = this.getNotificationDismissTime(alert.priority);
    if (dismissTime > 0) {
      setTimeout(() => {
        if (document.getElementById(`alert-toast-${alert.id}`)) {
          toast.style.animation = 'slideOutRight 0.3s cubic-bezier(0.16, 1, 0.3, 1)';
          setTimeout(() => toast.remove(), 300);
        }
      }, dismissTime);
    }

    // Add CSS animation if not already present
    if (!document.getElementById('alert-toast-animations')) {
      const style = document.createElement('style');
      style.id = 'alert-toast-animations';
      style.textContent = `
        @keyframes slideInRight {
          from { transform: translateX(100%); opacity: 0; }
          to { transform: translateX(0); opacity: 1; }
        }
        @keyframes slideOutRight {
          from { transform: translateX(0); opacity: 1; }
          to { transform: translateX(100%); opacity: 0; }
        }
      `;
      document.head.appendChild(style);
    }
  }

  // Get notification dismiss time based on priority
  private getNotificationDismissTime(priority: AlertPriority): number {
    switch (priority) {
      case AlertPriority.Critical: return 0; // Never auto-dismiss
      case AlertPriority.High: return 10000; // 10 seconds
      case AlertPriority.Medium: return 6000; // 6 seconds
      case AlertPriority.Low: return 4000; // 4 seconds
      default: return 6000;
    }
  }

  // Get priority emoji for email subjects
  private getPriorityEmoji(priority: AlertPriority): string {
    switch (priority) {
      case AlertPriority.Critical: return "🚨";
      case AlertPriority.High: return "⚠️";
      case AlertPriority.Medium: return "📢";
      case AlertPriority.Low: return "ℹ️";
      default: return "📢";
    }
  }

  // Public method to show success notifications for actions
  public showSuccessToast(message: string, title: string = "Success"): void {
    const alert: Partial<IAlertItem> = {
      id: Date.now().toString(),
      title: title,
      description: message,
      priority: AlertPriority.Low
    };
    this.createModernNotificationToast(alert as IAlertItem, { icon: "✅", color: "#107c10" });
  }

  // Public method to show error notifications for actions
  public showErrorToast(message: string, title: string = "Error"): void {
    const alert: Partial<IAlertItem> = {
      id: Date.now().toString(),
      title: title,
      description: message,
      priority: AlertPriority.High
    };
    this.createModernNotificationToast(alert as IAlertItem, { icon: "❌", color: "#d13438" });
  }

  // Public method to show warning notifications for actions
  public showWarningToast(message: string, title: string = "Warning"): void {
    const alert: Partial<IAlertItem> = {
      id: Date.now().toString(),
      title: title,
      description: message,
      priority: AlertPriority.Medium
    };
    this.createModernNotificationToast(alert as IAlertItem, { icon: "⚠️", color: "#f7630c" });
  }
}

export default NotificationService;