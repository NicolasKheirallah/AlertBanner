import * as React from "react";
import {
  DefaultButton,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
} from "@fluentui/react";
import { AppsListDetail24Regular } from "@fluentui/react-icons";
import { CopilotService, IGovernanceResult } from "../Services/CopilotService";
import { logger } from "../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import styles from "./CopilotGovernanceControl.module.scss";

export interface ICopilotGovernanceControlProps {
  copilotService: CopilotService;
  textToAnalyze: string;
  onError: (error: string) => void;
  disabled?: boolean;
}

const statusToBadgeClass = (status: IGovernanceResult["status"]): string => {
  switch (status) {
    case "green":
      return styles.statusGreen;
    case "yellow":
      return styles.statusYellow;
    case "red":
      return styles.statusRed;
    default:
      return styles.statusYellow;
  }
};

const statusToLabel = (status: IGovernanceResult["status"]): string => {
  switch (status) {
    case "green":
      return strings.CopilotGovernanceStatusGreen;
    case "yellow":
      return strings.CopilotGovernanceStatusYellow;
    case "red":
      return strings.CopilotGovernanceStatusRed;
    default:
      return status;
  }
};

export const CopilotGovernanceControl: React.FC<
  ICopilotGovernanceControlProps
> = ({ copilotService, textToAnalyze, onError, disabled }) => {
  const [isChecking, setIsChecking] = React.useState(false);
  const [isOpen, setIsOpen] = React.useState(false);
  const [result, setResult] = React.useState<IGovernanceResult | null>(null);

  const handleCheck = async (): Promise<void> => {
    if (!textToAnalyze) return;

    setIsChecking(true);
    try {
      const response = await copilotService.analyzeSentiment(textToAnalyze);

      if (response.isError) {
        if (!response.isCancelled) {
          onError(response.errorMessage || strings.CopilotAnalyzeFailed);
        }
      } else {
        const parsed = copilotService.parseGovernanceResult(response.content);
        setResult(parsed);
        setIsOpen(true);
      }
    } catch (error) {
      logger.error(
        "CopilotGovernanceControl",
        "Governance check failed",
        error,
      );
      onError(strings.CopilotUnexpectedError);
    } finally {
      setIsChecking(false);
    }
  };

  return (
    <>
      <div className={styles.governanceButton}>
        <DefaultButton
          onRenderIcon={() => <AppsListDetail24Regular />}
          onClick={handleCheck}
          disabled={disabled || isChecking || !textToAnalyze}
          className={styles.inlineGhostButton}
        >
          {isChecking
            ? strings.CopilotCheckingLabel
            : strings.CopilotGovernanceButton}
        </DefaultButton>
      </div>

      <Dialog
        hidden={!isOpen}
        onDismiss={() => setIsOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.CopilotGovernanceAnalysisTitle,
        }}
        modalProps={{
          isBlocking: false,
        }}
      >
        {result ? (
          <div>
            <span className={`${styles.statusBadge} ${statusToBadgeClass(result.status)}`}>
              {statusToLabel(result.status)}
            </span>

            {result.issues.length > 0 && (
              <ul className={styles.issueList}>
                {result.issues.map((issue, index) => (
                  <li key={index}>{issue}</li>
                ))}
              </ul>
            )}

            <div className={styles.summarySection}>
              <h4>{strings.CopilotGovernanceSummaryLabel}</h4>
              <p>{result.rawContent}</p>
            </div>
          </div>
        ) : (
          <div>{strings.CopilotNoIssuesFound}</div>
        )}
        <DialogFooter>
          <PrimaryButton onClick={() => setIsOpen(false)}>
            {strings.Close}
          </PrimaryButton>
        </DialogFooter>
      </Dialog>
    </>
  );
};
