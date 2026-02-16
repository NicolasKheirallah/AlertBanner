import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { logger } from './LoggerService';
import { UI_CONFIG, NOTIFICATION_STYLES } from '../Utils/AppConstants';

export enum NotificationType {
  Success = 'success',
  Warning = 'warning',
  Error = 'error',
  Info = 'info'
}

export interface INotificationAction {
  text: string;
  onClick: () => void;
}

export interface INotificationOptions {
  title?: string;
  message: string;
  type: NotificationType;
  duration?: number; // in milliseconds, 0 for persistent
  actions?: INotificationAction[];
  dismissible?: boolean;
}

/**
 * Service for displaying user-friendly notifications instead of browser alerts
 */
export class NotificationService {
  private static instance: NotificationService;
  private static readonly DIALOG_HOST_SELECTOR =
    '[data-alert-banner-dialog-notification-host="true"]';
  private static readonly NOTIFICATION_CONTAINER_SELECTOR =
    '[data-alert-banner-notification-container="true"]';
  private context: ApplicationCustomizerContext;

  constructor(context: ApplicationCustomizerContext) {
    this.context = context;
  }

  public static getInstance(context?: ApplicationCustomizerContext): NotificationService {
    if (!NotificationService.instance && context) {
      NotificationService.instance = new NotificationService(context);
    }
    return NotificationService.instance;
  }

  /**
   * Show a success notification
   */
  public showSuccess(message: string, title?: string, actions?: INotificationAction[]): void {
    this.showNotification({
      type: NotificationType.Success,
      title: title || 'Success',
      message,
      duration: UI_CONFIG.TOAST_DURATION_MEDIUM_MS,
      actions,
      dismissible: true
    });
  }

  /**
   * Show a warning notification
   */
  public showWarning(message: string, title?: string, actions?: INotificationAction[]): void {
    this.showNotification({
      type: NotificationType.Warning,
      title: title || 'Warning',
      message,
      duration: UI_CONFIG.TOAST_DURATION_LONG_MS,
      actions,
      dismissible: true
    });
  }

  /**
   * Show an error notification
   */
  public showError(message: string, title?: string, actions?: INotificationAction[]): void {
    this.showNotification({
      type: NotificationType.Error,
      title: title || 'Error',
      message,
      duration: 0, // Persistent
      actions,
      dismissible: true
    });
  }

  /**
   * Show an info notification
   */
  public showInfo(message: string, title?: string, actions?: INotificationAction[]): void {
    this.showNotification({
      type: NotificationType.Info,
      title: title || 'Information',
      message,
      duration: UI_CONFIG.TOAST_DURATION_MEDIUM_MS + 1000, 
      actions,
      dismissible: true
    });
  }

  /**
   * Show a notification with full options
   */
  private showNotification(options: INotificationOptions): void {
    try {
      // Use custom notification system for better control
      this.showCustomNotification(options);
    } catch (error) {
      logger.error('NotificationService', 'Failed to show notification, falling back to console', error);
      logger.info('NotificationService', `[${options.type.toUpperCase()}] ${options.title}: ${options.message}`);
    }
  }

  /**
   * Show custom notification using DOM manipulation
   */
  private showCustomNotification(options: INotificationOptions): void {
    const container = this.getOrCreateNotificationContainer();
    const notification = this.createNotificationElement(options);
    
    container.appendChild(notification);

    // Auto-dismiss if duration is set
    if (options.duration && options.duration > 0) {
      setTimeout(() => {
        this.removeNotification(notification);
      }, options.duration);
    }

    // Add click handlers for actions
    if (options.actions) {
      options.actions.forEach((action, index) => {
        const actionButton = notification.querySelector(`[data-action="${index}"]`);
        if (actionButton) {
          actionButton.addEventListener('click', () => {
            action.onClick();
            this.removeNotification(notification);
          });
        }
      });
    }

    // Add dismiss handler
    if (options.dismissible) {
      const dismissButton = notification.querySelector('[data-dismiss]');
      if (dismissButton) {
        dismissButton.addEventListener('click', () => {
          this.removeNotification(notification);
        });
      }
    }
  }

  /**
   * Get or create the notification container
   */
  private getOrCreateNotificationContainer(): HTMLElement {
    const hosts = Array.from(
      document.querySelectorAll<HTMLElement>(
        NotificationService.DIALOG_HOST_SELECTOR,
      ),
    );
    const host = hosts[hosts.length - 1];
    if (!host) {
      throw new Error('No dialog notification host is available.');
    }

    let container = host.querySelector<HTMLElement>(
      NotificationService.NOTIFICATION_CONTAINER_SELECTOR,
    );
    if (!container) {
      container = document.createElement('div');
      container.setAttribute('data-alert-banner-notification-container', 'true');
      container.setAttribute('role', 'region');
      container.setAttribute('aria-live', 'polite');
      container.setAttribute('aria-label', 'Alert banner notifications');
      container.style.cssText = `
        display: flex;
        flex-direction: column;
        gap: 12px;
      `;
      host.appendChild(container);
    }

    return container;
  }

  /**
   * Create notification DOM element
   */
  private createNotificationElement(options: INotificationOptions): HTMLElement {
    const notification = document.createElement('div');
    notification.setAttribute(
      'role',
      options.type === NotificationType.Error || options.type === NotificationType.Warning
        ? 'alert'
        : 'status'
    );
    notification.style.cssText = `
      background: ${this.getBackgroundColor(options.type)};
      color: ${this.getTextColor(options.type)};
      border-left: 4px solid ${this.getBorderColor(options.type)};
      border-radius: 6px;
      padding: 16px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      pointer-events: auto;
      font-family: 'Segoe UI', sans-serif;
      font-size: 14px;
      animation: slideInRight 0.3s ease-out;
    `;

    const wrapper = document.createElement('div');
    wrapper.style.position = 'relative';
    notification.appendChild(wrapper);

    if (options.dismissible) {
      const dismiss = document.createElement('button');
      dismiss.type = 'button';
      dismiss.dataset.dismiss = '';
      dismiss.textContent = '×';
      dismiss.style.cssText = `
        position: absolute;
        top: 8px;
        right: 8px;
        background: transparent;
        border: none;
        color: inherit;
        cursor: pointer;
        font-size: 16px;
        opacity: 0.7;
      `;
      wrapper.appendChild(dismiss);
    }

    const contentRow = document.createElement('div');
    contentRow.style.cssText = 'display: flex; align-items: flex-start; gap: 8px;';
    wrapper.appendChild(contentRow);

    const iconContainer = document.createElement('span');
    iconContainer.style.cssText = 'font-size: 16px; flex-shrink: 0;';
    iconContainer.textContent = this.getIcon(options.type);
    contentRow.appendChild(iconContainer);

    const textContainer = document.createElement('div');
    textContainer.style.cssText = 'flex: 1;';
    contentRow.appendChild(textContainer);

    if (options.title) {
      const title = document.createElement('div');
      title.style.cssText = 'font-weight: 600; margin-bottom: 4px;';
      title.textContent = options.title;
      textContainer.appendChild(title);
    }

    const message = document.createElement('div');
    message.style.cssText = 'line-height: 1.4;';
    message.textContent = options.message;
    textContainer.appendChild(message);

    if (options.actions && options.actions.length > 0) {
      const actionsRow = document.createElement('div');
      actionsRow.style.cssText = 'margin-top: 8px;';
      textContainer.appendChild(actionsRow);

      options.actions.forEach((action, index) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.dataset.action = index.toString();
        button.textContent = action.text;
        button.style.cssText = `
          background: transparent;
          border: 1px solid currentColor;
          color: inherit;
          padding: 4px 12px;
          border-radius: 4px;
          margin-left: ${index === 0 ? '0' : '8px'};
          cursor: pointer;
          font-size: 12px;
        `;
        actionsRow.appendChild(button);
      });
    }

    // Add animation styles to head if not already present
    this.addAnimationStyles();

    return notification;
  }

  /**
   * Remove notification element
   */
  private removeNotification(notification: HTMLElement): void {
    notification.style.animation = 'slideOutRight 0.3s ease-in';
    setTimeout(() => {
      if (notification.parentNode) {
        notification.parentNode.removeChild(notification);
      }
    }, UI_CONFIG.FADE_OUT_DURATION_MS + 100);
  }

  /**
   * Get background color for notification type
   */
  private getBackgroundColor(type: NotificationType): string {
    switch (type) {
      case NotificationType.Success:
        return NOTIFICATION_STYLES.SUCCESS.backgroundColor;
      case NotificationType.Warning:
        return NOTIFICATION_STYLES.WARNING.backgroundColor;
      case NotificationType.Error:
        return NOTIFICATION_STYLES.ERROR.backgroundColor;
      case NotificationType.Info:
      default:
        return NOTIFICATION_STYLES.INFO.backgroundColor;
    }
  }

  /**
   * Get text color for notification type
   */
  private getTextColor(type: NotificationType): string {
    switch (type) {
      case NotificationType.Success:
        return NOTIFICATION_STYLES.SUCCESS.textColor;
      case NotificationType.Warning:
        return NOTIFICATION_STYLES.WARNING.textColor;
      case NotificationType.Error:
        return NOTIFICATION_STYLES.ERROR.textColor;
      case NotificationType.Info:
      default:
        return NOTIFICATION_STYLES.INFO.textColor;
    }
  }

  /**
   * Get border color for notification type
   */
  private getBorderColor(type: NotificationType): string {
    switch (type) {
      case NotificationType.Success:
        return NOTIFICATION_STYLES.SUCCESS.borderColor;
      case NotificationType.Warning:
        return NOTIFICATION_STYLES.WARNING.borderColor;
      case NotificationType.Error:
        return NOTIFICATION_STYLES.ERROR.borderColor;
      case NotificationType.Info:
      default:
        return NOTIFICATION_STYLES.INFO.borderColor;
    }
  }

  /**
   * Get icon for notification type
   */
  private getIcon(type: NotificationType): string {
    switch (type) {
      case NotificationType.Success:
        return NOTIFICATION_STYLES.SUCCESS.icon;
      case NotificationType.Warning:
        return NOTIFICATION_STYLES.WARNING.icon;
      case NotificationType.Error:
        return NOTIFICATION_STYLES.ERROR.icon;
      case NotificationType.Info:
      default:
        return NOTIFICATION_STYLES.INFO.icon;
    }
  }

  /**
   * Add CSS animations to document head
   */
  private addAnimationStyles(): void {
    if (document.getElementById('alert-banner-notification-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'alert-banner-notification-styles';
    style.textContent = `
      @keyframes slideInRight {
        from {
          transform: translateX(400px);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      
      @keyframes slideOutRight {
        from {
          transform: translateX(0);
          opacity: 1;
        }
        to {
          transform: translateX(400px);
          opacity: 0;
        }
      }
    `;
    document.head.appendChild(style);
  }

  /**
   * Clear all notifications
   */
  public clearAll(): void {
    const containers = Array.from(
      document.querySelectorAll<HTMLElement>(
        NotificationService.NOTIFICATION_CONTAINER_SELECTOR,
      ),
    );
    containers.forEach((container) => {
      container.innerHTML = '';
    });
  }
}
