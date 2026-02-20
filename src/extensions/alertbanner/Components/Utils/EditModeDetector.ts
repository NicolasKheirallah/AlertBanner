import { logger } from '../Services/LoggerService';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';

export class EditModeDetector {
  public static isPageInEditMode(context?: ApplicationCustomizerContext): boolean {
    try {
      if (context) {
        const pageContext = context.pageContext;
        if (pageContext && pageContext.list && pageContext.list.permissions) {
          const hasEditPermission = pageContext.list.permissions.hasPermission(SPPermission.editListItems as any);
          if (!hasEditPermission) {
            logger.debug('EditModeDetector', 'User does not have edit permissions - not in edit mode');
          }
        }

        const legacy = (pageContext.legacyPageContext || {}) as any;
        if (typeof legacy.isPageInEditMode === 'boolean') {
          logger.debug('EditModeDetector', 'Edit mode detected via legacyPageContext.isPageInEditMode');
          return legacy.isPageInEditMode;
        }
        if (typeof legacy.isEditMode === 'boolean') {
          logger.debug('EditModeDetector', 'Edit mode detected via legacyPageContext.isEditMode');
          return legacy.isEditMode;
        }
        if (legacy.formContext?.displayMode) {
          const displayMode = String(legacy.formContext.displayMode).toLowerCase();
          if (displayMode === 'edit') {
            logger.debug('EditModeDetector', 'Edit mode detected via legacyPageContext.formContext.displayMode');
            return true;
          }
        }
        if (legacy.pageMode) {
          const pageMode = String(legacy.pageMode).toLowerCase();
          if (pageMode === 'edit') {
            logger.debug('EditModeDetector', 'Edit mode detected via legacyPageContext.pageMode');
            return true;
          }
        }
      }

      logger.debug('EditModeDetector', 'Checking edit mode', {
        url: window.location.href,
        bodyClasses: document.body.className
      });

      const urlParams = new URLSearchParams(window.location.search);
      const modeParam = urlParams.get('Mode') || urlParams.get('mode');
      const displayModeParam = urlParams.get('displaymode') || urlParams.get('DisplayMode');

      if ((modeParam && modeParam.toLowerCase() === 'edit') || (displayModeParam && displayModeParam.toLowerCase() === 'edit')) {
        logger.debug('EditModeDetector', 'Edit mode detected via URL params', { modeParam, displayModeParam });
        return true;
      }

      const bodyClasses = document.body.className;
      const editModeClasses = [
        'SPPageInEditMode',
        'ms-webpart-chrome-editing',
        'CanvasComponent-inEditMode',
        'od-EditMode',
        'SPCanvas--editing',
        'sp-edit-mode',
        'sp-editing',
        'ms-EditMode',
        'is-edit-mode'
      ];

      const foundEditClass = editModeClasses.find(cls => bodyClasses.includes(cls));
      if (foundEditClass) {
        logger.debug('EditModeDetector', 'Edit mode detected via body class', { className: foundEditClass });
        return true;
      }

      const spInfo = (window as any)._spPageContextInfo;
      if (spInfo) {
        if (spInfo.isPageInEditMode === true || spInfo.isEditMode === true) {
          logger.debug('EditModeDetector', 'Edit mode detected via _spPageContextInfo');
          return true;
        }
        if (typeof spInfo.pageMode === 'string' && spInfo.pageMode.toLowerCase() === 'edit') {
          logger.debug('EditModeDetector', 'Edit mode detected via _spPageContextInfo.pageMode');
          return true;
        }
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


      const commandBarSelectors = [
        '[data-automation-id="pageCommandBarRegion"]',
        '[data-automation-id="pageCommandBar"]',
        '[data-automation-id="PageCommandBar"]',
        '[data-automation-id="CommandBar"]'
      ];
      for (const selector of commandBarSelectors) {
        const editButtons = document.querySelector(selector);
        if (!editButtons) continue;
        const saveButton = editButtons.querySelector('button[title*="Save"]') ||
                          editButtons.querySelector('button[aria-label*="Save"]') ||
                          editButtons.querySelector('button[name*="Save"]');
        const publishButton = editButtons.querySelector('button[title*="Publish"]') ||
                             editButtons.querySelector('button[aria-label*="Publish"]') ||
                             editButtons.querySelector('button[name*="Publish"]');
        const republishButton = editButtons.querySelector('button[title*="Republish"]') ||
                               editButtons.querySelector('button[aria-label*="Republish"]');
        const discardButton = editButtons.querySelector('button[title*="Discard"]') ||
                             editButtons.querySelector('button[aria-label*="Discard"]') ||
                             editButtons.querySelector('button[title*="Cancel"]') ||
                             editButtons.querySelector('button[aria-label*="Cancel"]');
        const checkInButton = editButtons.querySelector('button[title*="Check in"]') ||
                             editButtons.querySelector('button[aria-label*="Check in"]');
        const submitButton = editButtons.querySelector('button[title*="Submit"]') ||
                             editButtons.querySelector('button[aria-label*="Submit"]');

        if (saveButton || publishButton || republishButton || discardButton || checkInButton || submitButton) {
          logger.debug('EditModeDetector', 'Edit mode detected via page command bar buttons');
          return true;
        }
      }

      const contentEditable = document.querySelector('[contenteditable="true"][data-sp-rte]') ||
                              document.querySelector('[contenteditable="true"][role="textbox"]');
      if (contentEditable) {
        logger.debug('EditModeDetector', 'Edit mode detected via contenteditable canvas');
        return true;
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
