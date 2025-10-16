import { logger } from '../Services/LoggerService';

export class EditModeDetector {
  public static isPageInEditMode(): boolean {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      if (urlParams.get('Mode') === 'Edit' || urlParams.get('displaymode') === 'edit') {
        return true;
      }

      const bodyClasses = document.body.className;
      if (bodyClasses.includes('SPPageInEditMode') ||
          bodyClasses.includes('ms-webpart-chrome-editing') ||
          bodyClasses.includes('CanvasComponent-inEditMode')) {
        return true;
      }

      const spPageDiv = document.querySelector('[data-sp-feature-tag="Site Pages Editing"]');
      if (spPageDiv) {
        return true;
      }

      const canvasEditingElements = document.querySelectorAll(
        '.CanvasComponent[data-sp-canvascontrol]'
      );
      if (canvasEditingElements.length > 0) {
        for (let i = 0; i < canvasEditingElements.length; i++) {
          const element = canvasEditingElements[i] as HTMLElement;
          if (element.style.outline || element.dataset.spCanvascontrol === 'editing') {
            return true;
          }
        }
      }

      const maintenanceIndicators = document.querySelectorAll(
        '.ms-webpartPage-root[data-automation-id="pageHeader"]'
      );
      if (maintenanceIndicators.length > 0) {
        return true;
      }

      const editButtons = document.querySelector('[data-automation-id="pageCommandBarRegion"]');
      if (editButtons) {
        const saveButton = editButtons.querySelector('button[title*="Save"]') ||
                          editButtons.querySelector('button[aria-label*="Save"]');
        const publishButton = editButtons.querySelector('button[title*="Publish"]') ||
                             editButtons.querySelector('button[aria-label*="Publish"]');

        if (saveButton || publishButton) {
          return true;
        }
      }

      if (typeof (window as any)._spPageContextInfo !== 'undefined') {
        const webUIVersion = (window as any)._spPageContextInfo.webUIVersion;
        if (webUIVersion && webUIVersion === 15) {
          if ((window as any).MSOLayout_InDesignMode || (window as any).g_disableCheckoutInEditMode === false) {
            return true;
          }
        }
      }

      return false;
    } catch (error) {
      logger.warn('EditModeDetector', 'Error detecting edit mode', error);
      return false;
    }
  }

  public static onEditModeChange(callback: (isEditMode: boolean) => void): () => void {
    let currentEditMode = EditModeDetector.isPageInEditMode();

    const checkForChanges = (): void => {
      const newEditMode = EditModeDetector.isPageInEditMode();
      if (newEditMode !== currentEditMode) {
        currentEditMode = newEditMode;
        callback(newEditMode);
      }
    };

    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    history.pushState = function(...args) {
      originalPushState.apply(history, args);
      setTimeout(checkForChanges, 100);
    };

    history.replaceState = function(...args) {
      originalReplaceState.apply(history, args);
      setTimeout(checkForChanges, 100);
    };

    const popstateListener = (): void => {
      setTimeout(checkForChanges, 100);
    };
    window.addEventListener('popstate', popstateListener);

    const observer = new MutationObserver(() => {
      checkForChanges();
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['class', 'data-sp-feature-tag', 'data-automation-id']
    });

    setTimeout(checkForChanges, 1000);

    return () => {
      history.pushState = originalPushState;
      history.replaceState = originalReplaceState;
      window.removeEventListener('popstate', popstateListener);
      observer.disconnect();
    };
  }
}