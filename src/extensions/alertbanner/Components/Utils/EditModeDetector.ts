import { logger } from '../Services/LoggerService';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';

export class EditModeDetector {
  public static isPageInEditMode(context?: ApplicationCustomizerContext): boolean {
    try {
      // First check if user has edit permissions
      if (context) {
        const pageContext = context.pageContext;
        // Check if user has edit rights - if not, definitely not in edit mode
        if (pageContext && pageContext.list && pageContext.list.permissions) {
          const hasEditPermission = pageContext.list.permissions.hasPermission(SPPermission.editListItems);
          if (!hasEditPermission) {
            logger.debug('EditModeDetector', 'User does not have edit permissions - not in edit mode');
            return false;
          }
        }
      }

      // Log current URL and body classes for debugging
      logger.debug('EditModeDetector', 'Checking edit mode', {
        url: window.location.href,
        bodyClasses: document.body.className
      });

      const urlParams = new URLSearchParams(window.location.search);
      const modeParam = urlParams.get('Mode');
      const displayModeParam = urlParams.get('displaymode');

      if (modeParam === 'Edit' || displayModeParam === 'edit') {
        logger.debug('EditModeDetector', 'Edit mode detected via URL params', { modeParam, displayModeParam });
        return true;
      }

      const bodyClasses = document.body.className;
      const editModeClasses = [
        'SPPageInEditMode',
        'ms-webpart-chrome-editing',
        'CanvasComponent-inEditMode',
        'od-EditMode',
        'SPCanvas--editing'
      ];

      const foundEditClass = editModeClasses.find(cls => bodyClasses.includes(cls));
      if (foundEditClass) {
        logger.debug('EditModeDetector', 'Edit mode detected via body class', { className: foundEditClass });
        return true;
      }

      const spPageDiv = document.querySelector('[data-sp-feature-tag="Site Pages Editing"]');
      if (spPageDiv) {
        logger.debug('EditModeDetector', 'Edit mode detected via Site Pages Editing feature tag');
        return true;
      }

      const canvasEditingElements = document.querySelectorAll(
        '.CanvasComponent[data-sp-canvascontrol]'
      );
      if (canvasEditingElements.length > 0) {
        for (let i = 0; i < canvasEditingElements.length; i++) {
          const element = canvasEditingElements[i] as HTMLElement;
          if (element.style.outline || element.dataset.spCanvascontrol === 'editing') {
            logger.debug('EditModeDetector', 'Edit mode detected via canvas control');
            return true;
          }
        }
      }

      const editControls = document.querySelector('[data-automation-id="editControls"]');
      if (editControls) {
        logger.debug('EditModeDetector', 'Edit mode detected via editControls');
        return true;
      }


      const editButtons = document.querySelector('[data-automation-id="pageCommandBarRegion"]');
      if (editButtons) {
        const saveButton = editButtons.querySelector('button[title*="Save"]') ||
                          editButtons.querySelector('button[aria-label*="Save"]') ||
                          editButtons.querySelector('button[name*="Save"]');
        const publishButton = editButtons.querySelector('button[title*="Publish"]') ||
                             editButtons.querySelector('button[aria-label*="Publish"]') ||
                             editButtons.querySelector('button[name*="Publish"]');
        const discardButton = editButtons.querySelector('button[title*="Discard"]') ||
                             editButtons.querySelector('button[aria-label*="Discard"]');

        if (saveButton || publishButton || discardButton) {
          logger.debug('EditModeDetector', 'Edit mode detected via page command bar buttons');
          return true;
        }
      }


      if (typeof (window as any)._spPageContextInfo !== 'undefined') {
        const webUIVersion = (window as any)._spPageContextInfo.webUIVersion;
        if (webUIVersion && webUIVersion === 15) {
          if ((window as any).MSOLayout_InDesignMode || (window as any).g_disableCheckoutInEditMode === false) {
            logger.debug('EditModeDetector', 'Edit mode detected via classic SharePoint mode');
            return true;
          }
        }
      }

      if (typeof (window as any).SP !== 'undefined' &&
          (window as any).SP?.Ribbon?.PageState?.Handlers?.isInEditMode) {
        const isInEdit = (window as any).SP.Ribbon.PageState.Handlers.isInEditMode();
        if (isInEdit) {
          logger.debug('EditModeDetector', 'Edit mode detected via SP Ribbon');
          return true;
        }
      }

      logger.debug('EditModeDetector', 'No edit mode detected - page is in view mode');
      return false;
    } catch (error) {
      logger.warn('EditModeDetector', 'Error detecting edit mode', error);
      return false;
    }
  }

}