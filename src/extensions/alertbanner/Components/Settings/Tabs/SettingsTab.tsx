import * as React from "react";
import { Add24Regular } from "@fluentui/react-icons";
import {
  SharePointButton,
  SharePointInput,
  SharePointToggle,
  SharePointSection
} from "../../UI/SharePointControls";
import { SharePointAlertService } from "../../Services/SharePointAlertService";
import { StorageService } from "../../Services/StorageService";
import styles from "../AlertSettings.module.scss";

export interface ISettingsData {
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
}

export interface ISettingsTabProps {
  settings: ISettingsData;
  setSettings: React.Dispatch<React.SetStateAction<ISettingsData>>;
  alertsListExists: boolean | null;
  setAlertsListExists: React.Dispatch<React.SetStateAction<boolean | null>>;
  alertTypesListExists: boolean | null;
  setAlertTypesListExists: React.Dispatch<React.SetStateAction<boolean | null>>;
  isCheckingLists: boolean;
  setIsCheckingLists: React.Dispatch<React.SetStateAction<boolean>>;
  isCreatingLists: boolean;
  setIsCreatingLists: React.Dispatch<React.SetStateAction<boolean>>;
  alertService: SharePointAlertService;
  onSettingsChange: (settings: ISettingsData) => void;
}

const SettingsTab: React.FC<ISettingsTabProps> = ({
  settings,
  setSettings,
  alertsListExists,
  setAlertsListExists,
  alertTypesListExists,
  setAlertTypesListExists,
  isCheckingLists,
  setIsCheckingLists,
  isCreatingLists,
  setIsCreatingLists,
  alertService,
  onSettingsChange
}) => {
  const storageService = React.useRef<StorageService>(StorageService.getInstance());
  const [carouselEnabled, setCarouselEnabled] = React.useState(false);
  const [carouselInterval, setCarouselInterval] = React.useState(5);

  // Load carousel settings from StorageService on mount
  React.useEffect(() => {
    const savedCarouselEnabled = storageService.current.getFromLocalStorage<boolean>('carouselEnabled');
    const savedCarouselInterval = storageService.current.getFromLocalStorage<number>('carouselInterval');
    
    if (savedCarouselEnabled !== null) {
      setCarouselEnabled(savedCarouselEnabled);
    }
    if (savedCarouselInterval && savedCarouselInterval >= 2000 && savedCarouselInterval <= 30000) {
      setCarouselInterval(savedCarouselInterval / 1000);
    }
  }, []);

  const handleCarouselEnabledChange = React.useCallback((checked: boolean) => {
    setCarouselEnabled(checked);
    storageService.current.saveToLocalStorage('carouselEnabled', checked);
    
    // Trigger a page refresh to apply changes
    setTimeout(() => window.location.reload(), 100);
  }, []);

  const handleCarouselIntervalChange = React.useCallback((value: string) => {
    const seconds = parseInt(value);
    if (seconds >= 2 && seconds <= 30) {
      setCarouselInterval(seconds);
      storageService.current.saveToLocalStorage('carouselInterval', seconds * 1000);
      
      // Trigger a page refresh to apply changes
      setTimeout(() => window.location.reload(), 100);
    }
  }, []);

  const handleSettingsChange = React.useCallback((newSettings: Partial<ISettingsData>) => {
    const updatedSettings = { ...settings, ...newSettings };
    setSettings(updatedSettings);
    onSettingsChange(updatedSettings);
  }, [settings, setSettings, onSettingsChange]);

  const checkListsExistence = React.useCallback(async () => {
    setIsCheckingLists(true);
    try {
      // Use the new detailed check method
      const listStatus = await alertService.checkListsNeeded();
      const currentSite = listStatus[0]; // Should be current site
      
      if (currentSite) {
        setAlertsListExists(currentSite.needsAlerts ? false : true);
        setAlertTypesListExists(currentSite.needsTypes ? false : true);
      } else {
        // Fallback to old method
        const [alertsTest, typesTest] = await Promise.allSettled([
          alertService.getAlerts(),
          alertService.getAlertTypes()
        ]);
        
        setAlertsListExists(alertsTest.status === 'fulfilled');
        setAlertTypesListExists(typesTest.status === 'fulfilled');
      }
    } catch (error) {
      console.error('Error checking lists:', error);
      // Fallback: assume lists don't exist if there's an error
      setAlertsListExists(false);
      setAlertTypesListExists(false);
    } finally {
      setIsCheckingLists(false);
    }
  }, [alertService, setAlertsListExists, setAlertTypesListExists, setIsCheckingLists]);

  const handleCreateLists = React.useCallback(async () => {
    setIsCreatingLists(true);
    try {
      // First check what's needed
      const listStatus = await alertService.checkListsNeeded();
      const currentSite = listStatus[0];
      
      if (!currentSite || (!currentSite.needsAlerts && !currentSite.needsTypes)) {
        alert('All required lists already exist on this site.');
        return;
      }
      
      // Initialize lists using the existing service method
      await alertService.initializeLists();
      
      // Re-check lists after creation
      await checkListsExistence();
      
      // Success message
      const createdLists = [];
      if (currentSite.needsAlerts) createdLists.push('Alerts');
      if (currentSite.needsTypes) createdLists.push('AlertBannerTypes');
      
      if (createdLists.length > 0) {
        alert(`Successfully created ${createdLists.join(' and ')} list${createdLists.length > 1 ? 's' : ''} on this site.`);
      }
    } catch (error) {
      console.error('Error creating lists:', error);
      const errorMsg = error.message || error.toString();
      
      if (errorMsg.includes('PERMISSION_DENIED')) {
        alert('Permission denied: You need site owner or full control permissions to create SharePoint lists.');
      } else {
        alert(`Failed to create some lists: ${errorMsg}`);
      }
    } finally {
      setIsCreatingLists(false);
    }
  }, [alertService, checkListsExistence, setIsCreatingLists]);

  // Check lists on mount
  React.useEffect(() => {
    checkListsExistence();
  }, [checkListsExistence]);

  return (
    <div className={styles.tabContent}>
      <SharePointSection title="Feature Settings">
        <div className={styles.settingsGrid}>
          <SharePointToggle
            label="Enable User Targeting"
            checked={settings.userTargetingEnabled}
            onChange={(checked) => handleSettingsChange({ userTargetingEnabled: checked })}
            description="Allow alerts to target specific users or groups based on SharePoint profiles and security groups"
          />

          <SharePointToggle
            label="Enable Browser Notifications"
            checked={settings.notificationsEnabled}
            onChange={(checked) => handleSettingsChange({ notificationsEnabled: checked })}
            description="Send native browser notifications for critical and high-priority alerts to ensure visibility"
          />

          <SharePointToggle
            label="Enable Rich Media Support"
            checked={settings.richMediaEnabled}
            onChange={(checked) => handleSettingsChange({ richMediaEnabled: checked })}
            description="Support images, videos, HTML content, and markdown formatting in alert descriptions"
          />
        </div>
      </SharePointSection>

      <SharePointSection title="Carousel Settings">
        <div className={styles.settingsGrid}>
          <SharePointToggle
            label="Enable Carousel Auto-Rotation"
            checked={carouselEnabled}
            onChange={handleCarouselEnabledChange}
            description="Automatically rotate between multiple alerts when more than one is displayed"
          />

          <SharePointInput
            label="Carousel Timer (seconds)"
            value={carouselInterval.toString()}
            onChange={handleCarouselIntervalChange}
            placeholder="5"
            type="text"
            description="Time in seconds between automatic alert transitions (2-30 seconds)"
            disabled={!carouselEnabled}
          />
        </div>
      </SharePointSection>

      {/* SharePoint Setup - Shows when lists are missing */}
      {(alertsListExists === false || alertTypesListExists === false) && (
        <SharePointSection title="SharePoint Setup Required">
          <div className={styles.settingsGrid}>
            <div className={styles.fullWidthColumn}>
              {isCheckingLists ? (
                <div className={styles.spinnerContainer}>
                  <div className={styles.spinner}></div>
                  Checking SharePoint lists...
                </div>
              ) : (
                <>
                  <p className={styles.infoText}>
                    The following lists are missing on this site and need to be created:
                  </p>
                  <div className={styles.infoText}>
                    <strong>Current Site:</strong> {window.location.href.split('/')[2]}
                  </div>
                  <ul className={styles.infoText}>
                    {alertsListExists === false && (
                      <li><strong>Alerts</strong> - For storing alert content on this site</li>
                    )}
                    {alertTypesListExists === false && (
                      <li><strong>AlertBannerTypes</strong> - For alert styling configurations (can be shared across sites)</li>
                    )}
                  </ul>
                  
                  <div className={styles.actionButtonsRow}>
                    <SharePointButton
                      variant="primary"
                      icon={<Add24Regular />}
                      onClick={handleCreateLists}
                      disabled={isCreatingLists}
                    >
                      {isCreatingLists ? 'Creating Lists...' : 'Create Missing Lists'}
                    </SharePointButton>
                    
                    <div className={styles.helpText}>
                      Creates only the missing lists on the current site.
                    </div>
                  </div>

                  {isCreatingLists && (
                    <div className={styles.creatingProgress}>
                      <div className={styles.spinnerContainer}>
                        <div className={styles.spinner}></div>
                        Creating SharePoint lists... This may take a few moments.
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </SharePointSection>
      )}

      {/* Success message when lists exist */}
      {alertsListExists === true && alertTypesListExists === true && (
        <SharePointSection title="SharePoint Setup">
          <div className={styles.successContainer}>
            <div className={styles.successHeader}>
              <span className={styles.successIcon}>✅</span>
              <strong>Setup Complete</strong>
            </div>
            <p className={styles.successDescription}>
              All required SharePoint lists are properly configured and ready to use.
            </p>
          </div>
        </SharePointSection>
      )}

      <SharePointSection title="Storage Management">
        <div className={styles.settingsGrid}>
          <div className={styles.fullWidthColumn}>
            <p className={styles.storageManagement}>
              Manage local storage and cached data for the Alert Banner system.
            </p>
            <div className={styles.storageButtons}>
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  storageService.current.clearAllAlertData();
                  alert('Alert data cleared from local storage.');
                }}
              >
                Clear Alert Cache
              </SharePointButton>
              <SharePointButton
                variant="secondary"
                onClick={() => {
                  storageService.current.removeFromLocalStorage('carouselEnabled');
                  storageService.current.removeFromLocalStorage('carouselInterval');
                  setCarouselEnabled(false);
                  setCarouselInterval(5);
                  alert('Carousel settings reset to defaults.');
                }}
              >
                Reset Carousel Settings
              </SharePointButton>
            </div>
          </div>
        </div>
      </SharePointSection>
    </div>
  );
};

export default SettingsTab;