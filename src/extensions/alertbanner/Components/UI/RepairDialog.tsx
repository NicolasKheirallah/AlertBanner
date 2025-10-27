import * as React from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  List,
  Text,
  Icon,
  Stack,
  Separator
} from '@fluentui/react';
import { IRepairResult } from '../Services/SharePointAlertService';
import { SharePointAlertService } from '../Services/SharePointAlertService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import styles from './RepairDialog.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';

const localize = (key: keyof typeof strings | string, ...args: Array<string | number>): string => {
  const dictionary = strings as unknown as Record<string, string>;
  const value = dictionary[key as string] ?? key.toString();
  if (args.length === 0) {
    return value;
  }
  const formattedArgs = args.map(arg => arg.toString());
  return CoreText.format(value, ...formattedArgs);
};

export interface IRepairDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  onRepairComplete: (result: IRepairResult) => void;
  alertService: SharePointAlertService;
}

interface IRepairProgress {
  message: string;
  progress: number;
}

const RepairDialog: React.FC<IRepairDialogProps> = ({
  isOpen,
  onDismiss,
  onRepairComplete,
  alertService
}) => {
  const [repairProgress, setRepairProgress] = React.useState<IRepairProgress>({ message: '', progress: 0 });
  const [showConfirmation, setShowConfirmation] = React.useState(true);

  const { loading: isRepairing, data: repairResult, execute: executeRepair, reset } = useAsyncOperation(
    async () => {
      const siteId = alertService.getCurrentSiteId();
      return await alertService.repairAlertsList(
        siteId,
        (message: string, progress: number) => {
          setRepairProgress({ message, progress });
        }
      );
    },
    {
      onSuccess: (result) => onRepairComplete(result),
      onError: (error) => {
        const errorResult: IRepairResult = {
          success: false,
          message: `Repair failed: ${error.message}`,
          details: {
            columnsRemoved: [],
            columnsAdded: [],
            columnsUpdated: [],
            errors: [error.message],
            warnings: []
          }
        };
        onRepairComplete(errorResult);
      },
      logErrors: true
    }
  );

  const handleStartRepair = React.useCallback(async () => {
    setShowConfirmation(false);
    await executeRepair();
  }, [executeRepair]);

  const handleDismiss = React.useCallback(() => {
    if (!isRepairing) {
      setShowConfirmation(true);
      setRepairProgress({ message: '', progress: 0 });
      reset();
      onDismiss();
    }
  }, [isRepairing, onDismiss, reset]);

  const renderConfirmationContent = () => (
    <Stack tokens={{ childrenGap: 20 }}>
      <MessageBar messageBarType={MessageBarType.warning}>
        <strong>{strings.RepairDialogWarningMessage}</strong>
      </MessageBar>
      
      <div className={styles.confirmationContent}>
        <Text variant="medium">
          {strings.RepairDialogActionsIntro}
        </Text>
        
        <ul className={styles.repairActionsList}>
          <li>
            <Icon iconName="Delete" className={styles.removeIcon} />
            <Text>{strings.RepairDialogRemoveColumns}</Text>
          </li>
          <li>
            <Icon iconName="Add" className={styles.addIcon} />
            <Text>{strings.RepairDialogAddColumns}</Text>
          </li>
          <li>
            <Icon iconName="Refresh" className={styles.updateIcon} />
            <Text>{strings.RepairDialogUpdateColumns}</Text>
          </li>
          <li>
            <Icon iconName="Shield" className={styles.protectIcon} />
            <Text>{strings.RepairDialogProtectData}</Text>
          </li>
        </ul>

        <MessageBar messageBarType={MessageBarType.info}>
          <strong>{strings.RepairDialogSafeProcess}</strong> {strings.RepairDialogSafeProcessDescription}
        </MessageBar>
      </div>
    </Stack>
  );

  const renderProgressContent = () => (
    <Stack tokens={{ childrenGap: 15 }}>
      <Text variant="mediumPlus">{strings.RepairDialogProgressTitle}</Text>
      
      <ProgressIndicator
        percentComplete={repairProgress.progress / 100}
        description={repairProgress.message}
        className={styles.progressIndicator}
      />
      
      <Text variant="small" className={styles.progressText}>
        {strings.RepairDialogProgressDescription}
      </Text>
    </Stack>
  );

  const renderResultContent = () => {
    if (!repairResult) return null;

    const { success, message, details } = repairResult;
    
    return (
      <Stack tokens={{ childrenGap: 20 }}>
        <MessageBar 
          messageBarType={success ? MessageBarType.success : MessageBarType.error}
          className={styles.resultMessage}
        >
          <strong>{success ? strings.RepairDialogResultSuccessTitle : strings.RepairDialogResultFailureTitle}</strong>
          <br />
          {message}
        </MessageBar>

        {details.columnsRemoved.length > 0 && (
          <div className={styles.resultSection}>
            <Text variant="medium" className={styles.sectionTitle}>
              <Icon iconName="Delete" className={styles.removeIcon} />
              {CoreText.format(strings.RepairDialogRemovedColumnsTitle, details.columnsRemoved.length.toString())}
            </Text>
            <List
              items={details.columnsRemoved}
              onRenderCell={(item) => (
                <div className={styles.listItem}>
                  <Icon iconName="CheckMark" className={styles.successIcon} />
                  <Text variant="small">{item}</Text>
                </div>
              )}
            />
          </div>
        )}

        {details.columnsAdded.length > 0 && (
          <div className={styles.resultSection}>
            <Text variant="medium" className={styles.sectionTitle}>
              <Icon iconName="Add" className={styles.addIcon} />
              {CoreText.format(strings.RepairDialogAddedColumnsTitle, details.columnsAdded.length.toString())}
            </Text>
            <List
              items={details.columnsAdded}
              onRenderCell={(item) => (
                <div className={styles.listItem}>
                  <Icon iconName="CheckMark" className={styles.successIcon} />
                  <Text variant="small">{item}</Text>
                </div>
              )}
            />
          </div>
        )}

        {details.warnings.length > 0 && (
          <div className={styles.resultSection}>
            <Text variant="medium" className={styles.sectionTitle}>
              <Icon iconName="Warning" className={styles.warningIcon} />
              {CoreText.format(strings.RepairDialogWarningsTitle, details.warnings.length.toString())}
            </Text>
            <List
              items={details.warnings}
              onRenderCell={(item) => (
                <div className={styles.listItem}>
                  <Icon iconName="Warning" className={styles.warningIcon} />
                  <Text variant="small">{item}</Text>
                </div>
              )}
            />
          </div>
        )}

        {details.errors.length > 0 && (
          <div className={styles.resultSection}>
            <Text variant="medium" className={styles.sectionTitle}>
              <Icon iconName="Error" className={styles.errorIcon} />
              {CoreText.format(strings.RepairDialogErrorsTitle, details.errors.length.toString())}
            </Text>
            <List
              items={details.errors}
              onRenderCell={(item) => (
                <div className={styles.listItem}>
                  <Icon iconName="Error" className={styles.errorIcon} />
                  <Text variant="small">{item}</Text>
                </div>
              )}
            />
          </div>
        )}
      </Stack>
    );
  };

  const getDialogTitle = () => {
    if (repairResult) {
      return repairResult.success ? strings.RepairDialogDialogTitleSuccess : strings.RepairDialogDialogTitleIssues;
    }
    if (isRepairing) {
      return strings.RepairDialogDialogTitleProgress;
    }
    return strings.RepairDialogTitle;
  };

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={handleDismiss}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: getDialogTitle(),
        subText: !showConfirmation && !repairResult ? strings.RepairDialogPleaseWait : undefined
      }}
      modalProps={{
        isBlocking: isRepairing,
        dragOptions: isRepairing ? undefined : {
          moveMenuItemText: 'Move',
          closeMenuItemText: 'Close',
          menu: undefined
        }
      }}
      minWidth={600}
      maxWidth={800}
      className={styles.repairDialog}
    >
      <div className={styles.dialogContent}>
        {showConfirmation && renderConfirmationContent()}
        {isRepairing && renderProgressContent()}
        {repairResult && renderResultContent()}
      </div>

      <DialogFooter>
        {showConfirmation && (
          <>
            <PrimaryButton
              onClick={handleStartRepair}
              text={strings.RepairDialogStartButton}
              iconProps={{ iconName: 'Wrench' }}
              className={styles.primaryButton}
            />
            <DefaultButton
              onClick={handleDismiss}
              text={strings.RepairDialogCancelButton}
            />
          </>
        )}
        
        {isRepairing && (
          <DefaultButton
            disabled
            text={strings.RepairDialogInProgressButton}
            iconProps={{ iconName: 'ProgressLoopInner' }}
          />
        )}
        
        {repairResult && (
          <PrimaryButton
            onClick={handleDismiss}
            text={strings.RepairDialogCloseButton}
            iconProps={{ iconName: 'CheckMark' }}
          />
        )}
      </DialogFooter>
    </Dialog>
  );
};

export default RepairDialog;
