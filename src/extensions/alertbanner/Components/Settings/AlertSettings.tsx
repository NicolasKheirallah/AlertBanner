import * as React from "react";
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogActions,
  DialogBody,
  Button,
  Field,
  Switch,
  Textarea,
  Text,
  tokens,
  Divider
} from "@fluentui/react-components";
import { Settings24Regular, Dismiss24Regular } from "@fluentui/react-icons";
import { useLocalization } from "../Hooks/useLocalization";
import LanguageSelector from "../UI/LanguageSelector";
import ListManagement from "../UI/ListManagement";
import { SiteContextService } from "../Services/SiteContextService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "./AlertSettings.module.scss";

export interface IAlertSettingsProps {
  isInEditMode: boolean;
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
  context?: ApplicationCustomizerContext;
  graphClient?: MSGraphClientV3;
  onSettingsChange: (settings: ISettingsData) => void;
}

export interface ISettingsData {
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
}

const AlertSettings: React.FC<IAlertSettingsProps> = ({
  isInEditMode,
  alertTypesJson,
  userTargetingEnabled,
  notificationsEnabled,
  richMediaEnabled,
  context,
  graphClient,
  onSettingsChange
}) => {
  const { getString } = useLocalization();
  const [isOpen, setIsOpen] = React.useState(false);
  const [siteContextService, setSiteContextService] = React.useState<SiteContextService | null>(null);
  const [settings, setSettings] = React.useState<ISettingsData>({
    alertTypesJson,
    userTargetingEnabled,
    notificationsEnabled,
    richMediaEnabled
  });

  const [alertTypesText, setAlertTypesText] = React.useState(() => {
    try {
      return JSON.stringify(JSON.parse(alertTypesJson), null, 2);
    } catch {
      return alertTypesJson;
    }
  });

  // Initialize site context service when dialog opens
  React.useEffect(() => {
    if (isOpen && context && graphClient && !siteContextService) {
      const service = SiteContextService.getInstance(context, graphClient);
      service.initialize().then(() => {
        setSiteContextService(service);
      }).catch(error => {
        console.error('Failed to initialize site context service:', error);
      });
    }
  }, [isOpen, context, graphClient, siteContextService]);

  // Don't render anything if not in edit mode
  if (!isInEditMode) {
    return null;
  }

  const handleSave = () => {
    try {
      // Validate JSON before saving
      JSON.parse(alertTypesText);
      const updatedSettings: ISettingsData = {
        ...settings,
        alertTypesJson: alertTypesText
      };
      onSettingsChange(updatedSettings);
      setIsOpen(false);
    } catch (error) {
      alert(getString('InvalidJSONError'));
    }
  };

  const handleCancel = () => {
    // Reset to original values
    setSettings({
      alertTypesJson,
      userTargetingEnabled,
      notificationsEnabled,
      richMediaEnabled
    });
    setAlertTypesText(() => {
      try {
        return JSON.stringify(JSON.parse(alertTypesJson), null, 2);
      } catch {
        return alertTypesJson;
      }
    });
    setIsOpen(false);
  };

  return (
    <Dialog open={isOpen} onOpenChange={(_, data) => setIsOpen(data.open)}>
      <DialogTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          icon={<Settings24Regular />}
          onClick={() => setIsOpen(true)}
          aria-label={getString('AlertSettings')}
          title={getString('ConfigureAlertBannerSettings')}
          className={styles.settingsButton}
          style={{
            backgroundColor: tokens.colorNeutralBackground1,
            border: `1px solid ${tokens.colorNeutralStroke1}`,
            boxShadow: tokens.shadow4
          }}
        />
      </DialogTrigger>
      <DialogSurface style={{ minWidth: '600px', maxWidth: '800px' }}>
        <DialogBody>
          <DialogTitle
            action={
              <DialogTrigger action="close">
                <Button
                  appearance="subtle"
                  aria-label={getString('Close')}
                  icon={<Dismiss24Regular />}
                />
              </DialogTrigger>
            }
          >
            {getString('AlertSettingsTitle')}
          </DialogTitle>
          <DialogContent className={styles.settingsContent}
          >
            <Text>
              {getString('AlertSettingsDescription')}
            </Text>

            <Divider />

            <Field label={getString('Language')}>
              <LanguageSelector />
            </Field>

            <Divider />

            <Field label={getString('Features')}>
              <div className={styles.featureSection}>
                <Field label={getString('EnableUserTargeting')}>
                  <Switch
                    checked={settings.userTargetingEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, userTargetingEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    {getString('EnableUserTargetingDescription')}
                  </Text>
                </Field>

                <Field label={getString('EnableNotifications')}>
                  <Switch
                    checked={settings.notificationsEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, notificationsEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    {getString('EnableNotificationsDescription')}
                  </Text>
                </Field>

                <Field label={getString('EnableRichMedia')}>
                  <Switch
                    checked={settings.richMediaEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, richMediaEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    {getString('EnableRichMediaDescription')}
                  </Text>
                </Field>
              </div>
            </Field>

            <Divider />

            <Field label={getString('AlertTypesConfiguration')}>
              <Text size={200} style={{ marginBottom: tokens.spacingVerticalS }}>
                {getString('AlertTypesConfigurationDescription')}
              </Text>
              <Textarea
                value={alertTypesText}
                onChange={(_, data) => setAlertTypesText(data.value)}
                rows={15}
                className={styles.configTextarea}
                placeholder={getString('AlertTypesPlaceholder')}
              />
              <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                {getString('AlertTypesHelpText')}
              </Text>
            </Field>

            {siteContextService && (
              <>
                <Divider />
                <Text weight="semibold" size={300}>
                  {getString('AlertListsManagement') || 'Alert Lists Management'}
                </Text>
                <ListManagement 
                  siteContextService={siteContextService}
                  onListCreated={() => {
                    // Refresh the site context when a list is created
                    siteContextService.refresh();
                  }}
                />
              </>
            )}
          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={handleCancel}>
                {getString('Cancel')}
              </Button>
            </DialogTrigger>
            <Button appearance="primary" onClick={handleSave}>
              {getString('SaveSettings')}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default AlertSettings;