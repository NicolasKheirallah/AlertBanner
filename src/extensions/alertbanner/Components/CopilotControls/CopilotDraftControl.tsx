import * as React from "react";
import {
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Button,
  Textarea,
  Spinner,
  Dropdown,
  Option,
  Label,
  useId,
} from "@fluentui/react-components";
import { SparkleRegular } from "@fluentui/react-icons";
import { CopilotService, CopilotTone } from "../Services/CopilotService";
import { logger } from "../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import styles from "./CopilotDraftControl.module.scss";

export interface ICopilotDraftControlProps {
  copilotService: CopilotService;
  onDraftGenerated: (draft: string) => void;
  onError: (error: string) => void;
  disabled?: boolean;
}

export const CopilotDraftControl: React.FC<ICopilotDraftControlProps> = ({
  copilotService,
  onDraftGenerated,
  onError,
  disabled,
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [isDrafting, setIsDrafting] = React.useState(false);
  const [keywords, setKeywords] = React.useState("");
  const [tone, setTone] = React.useState<CopilotTone>("Professional");
  const dropdownId = useId("tone-dropdown");

  const handleGenerate = async (): Promise<void> => {
    if (!keywords.trim()) return;

    setIsDrafting(true);
    try {
      const response = await copilotService.generateDraft(keywords, tone);

      if (response.isError) {
        if (!response.isCancelled) {
          onError(response.errorMessage || strings.CopilotDraftGenerationFailed);
        }
      } else {
        onDraftGenerated(response.content);
        setIsOpen(false);
        setKeywords("");
      }
    } catch (error) {
      logger.error("CopilotDraftControl", "Draft generation failed", error);
      onError(strings.CopilotUnexpectedError);
    } finally {
      setIsDrafting(false);
    }
  };

  const handleCancel = (): void => {
    copilotService.cancelActiveOperation();
    setIsDrafting(false);
    setIsOpen(false);
  };

  return (
    <Popover open={isOpen} onOpenChange={(_, data) => setIsOpen(data.open)}>
      <PopoverTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          icon={<SparkleRegular />}
          disabled={disabled || isDrafting}
          size="small"
        >
          {strings.CopilotDraftButton}
        </Button>
      </PopoverTrigger>
      <PopoverSurface tabIndex={-1} className={styles.popoverContent}>
        <h3 className={styles.popoverTitle}>{strings.CopilotDraftTitle}</h3>

        <div className={styles.fieldGroup}>
          <Label htmlFor="copilot-keywords">
            {strings.CopilotKeywordsLabel}
          </Label>
          <Textarea
            id="copilot-keywords"
            placeholder={strings.CopilotKeywordsPlaceholder}
            value={keywords}
            onChange={(_, data) => setKeywords(data.value)}
            className={styles.fullWidth}
            rows={3}
          />
        </div>

        <div className={styles.toneGroup}>
          <Label htmlFor={dropdownId}>{strings.CopilotToneLabel}</Label>
          <Dropdown
            id={dropdownId}
            value={tone}
            selectedOptions={[tone]}
            onOptionSelect={(_, data) => {
              if (data.optionValue) {
                setTone(data.optionValue as CopilotTone);
              }
            }}
            className={styles.fullWidth}
          >
            <Option value="Professional">
              {strings.CopilotToneProfessional}
            </Option>
            <Option value="Urgent">{strings.CopilotToneUrgent}</Option>
            <Option value="Casual">{strings.CopilotToneCasual}</Option>
          </Dropdown>
        </div>

        <div className={styles.actions}>
          {isDrafting && (
            <Button appearance="subtle" onClick={handleCancel} size="small">
              {strings.Cancel}
            </Button>
          )}
          <Button
            appearance="primary"
            onClick={handleGenerate}
            disabled={!keywords.trim() || isDrafting}
            icon={isDrafting ? <Spinner size="tiny" /> : undefined}
          >
            {isDrafting
              ? strings.CopilotGeneratingLabel
              : strings.CopilotGenerateButton}
          </Button>
        </div>
      </PopoverSurface>
    </Popover>
  );
};
