import * as React from "react";
import {
  Button,
  Spinner,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Badge,
} from "@fluentui/react-components";
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

/**
 * Maps a governance status to a Fluent UI Badge color.
 */
const statusToBadgeColor = (
  status: IGovernanceResult["status"],
): "success" | "warning" | "danger" => {
  switch (status) {
    case "green":
      return "success";
    case "yellow":
      return "warning";
    case "red":
      return "danger";
    default:
      return "warning";
  }
};

/**
 * Maps a governance status to a display label.
 */
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
        onError(response.errorMessage || strings.CopilotAnalyzeFailed);
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
        <Button
          appearance="subtle"
          icon={<AppsListDetail24Regular />}
          onClick={handleCheck}
          disabled={disabled || isChecking || !textToAnalyze}
          size="small"
        >
          {isChecking
            ? strings.CopilotCheckingLabel
            : strings.CopilotGovernanceButton}
        </Button>
      </div>

      <Dialog open={isOpen} onOpenChange={(_, data) => setIsOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.CopilotGovernanceAnalysisTitle}</DialogTitle>
            <DialogContent>
              {result ? (
                <div>
                  <Badge
                    color={statusToBadgeColor(result.status)}
                    appearance="filled"
                    size="large"
                  >
                    {statusToLabel(result.status)}
                  </Badge>

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
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setIsOpen(false)}>
                {strings.Close}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};
