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
import styles from "./AlertSettings.module.scss";

export interface IAlertSettingsProps {
  isInEditMode: boolean;
  alertTypesJson: string;
  userTargetingEnabled: boolean;
  notificationsEnabled: boolean;
  richMediaEnabled: boolean;
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
  onSettingsChange
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
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
      alert('Invalid JSON format in Alert Types configuration. Please check your syntax.');
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
          aria-label="Alert Settings"
          title="Configure Alert Banner Settings"
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
                  aria-label="close"
                  icon={<Dismiss24Regular />}
                />
              </DialogTrigger>
            }
          >
            Alert Banner Settings
          </DialogTitle>
          <DialogContent className={styles.settingsContent}
          >
            <Text>
              Configure the alert banner settings. These changes will be applied site-wide.
            </Text>

            <Divider />

            <Field label="Features">
              <div className={styles.featureSection}>
                <Field label="Enable User Targeting">
                  <Switch
                    checked={settings.userTargetingEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, userTargetingEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    Allow alerts to target specific users or groups
                  </Text>
                </Field>

                <Field label="Enable Notifications">
                  <Switch
                    checked={settings.notificationsEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, notificationsEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    Send browser notifications for critical alerts
                  </Text>
                </Field>

                <Field label="Enable Rich Media">
                  <Switch
                    checked={settings.richMediaEnabled}
                    onChange={(_, data) =>
                      setSettings(prev => ({ ...prev, richMediaEnabled: data.checked }))
                    }
                  />
                  <Text size={200}>
                    Support images, videos, and rich content in alerts
                  </Text>
                </Field>
              </div>
            </Field>

            <Divider />

            <Field label="Alert Types Configuration">
              <Text size={200} style={{ marginBottom: tokens.spacingVerticalS }}>
                Configure the available alert types (JSON format):
              </Text>
              <Textarea
                value={alertTypesText}
                onChange={(_, data) => setAlertTypesText(data.value)}
                rows={15}
                className={styles.configTextarea}
                placeholder="Enter alert types JSON configuration..."
              />
              <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                Each alert type should have: name, iconName, backgroundColor, textColor, additionalStyles, and priorityStyles
              </Text>
            </Field>
          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={handleCancel}>
                Cancel
              </Button>
            </DialogTrigger>
            <Button appearance="primary" onClick={handleSave}>
              Save Settings
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default AlertSettings;