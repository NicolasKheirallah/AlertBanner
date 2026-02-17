import * as React from "react";
import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Spinner,
  SpinnerSize,
  Dropdown,
  IDropdownOption,
  Label,
  Callout,
  DirectionalHint,
} from "@fluentui/react";
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
  const triggerRef = React.useRef<HTMLDivElement | null>(null);
  const dropdownId = React.useMemo(
    () => `tone-dropdown-${Math.random().toString(36).slice(2, 10)}`,
    [],
  );

  const toneOptions = React.useMemo<IDropdownOption[]>(
    () => [
      { key: "Professional", text: strings.CopilotToneProfessional },
      { key: "Urgent", text: strings.CopilotToneUrgent },
      { key: "Casual", text: strings.CopilotToneCasual },
    ],
    [],
  );

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
    <>
      <div ref={triggerRef}>
        <DefaultButton
          onRenderIcon={() => <SparkleRegular />}
          disabled={disabled || isDrafting}
          onClick={() => setIsOpen((prev) => !prev)}
          className={styles.ghostButton}
        >
          {strings.CopilotDraftButton}
        </DefaultButton>
      </div>
      {isOpen && triggerRef.current && (
        <Callout
          target={triggerRef.current}
          onDismiss={() => setIsOpen(false)}
          directionalHint={DirectionalHint.bottomLeftEdge}
          setInitialFocus={true}
          isBeakVisible={false}
          gapSpace={8}
        >
          <div tabIndex={-1} className={styles.popoverContent}>
        <h3 className={styles.popoverTitle}>{strings.CopilotDraftTitle}</h3>

        <div className={styles.fieldGroup}>
          <Label htmlFor="copilot-keywords">
            {strings.CopilotKeywordsLabel}
          </Label>
          <TextField
            id="copilot-keywords"
            placeholder={strings.CopilotKeywordsPlaceholder}
            value={keywords}
            multiline
            rows={3}
            onChange={(_, value) => setKeywords(value || "")}
            className={styles.fullWidth}
          />
        </div>

        <div className={styles.toneGroup}>
          <Label htmlFor={dropdownId}>{strings.CopilotToneLabel}</Label>
          <Dropdown
            id={dropdownId}
            selectedKey={tone}
            options={toneOptions}
            onChange={(_, option) => {
              if (option?.key) {
                setTone(option.key as CopilotTone);
              }
            }}
            className={styles.fullWidth}
          />
        </div>

        <div className={styles.actions}>
          {isDrafting && (
            <DefaultButton
              onClick={handleCancel}
              className={`${styles.ghostButton} ${styles.ghostButtonCompact}`}
            >
              {strings.Cancel}
            </DefaultButton>
          )}
          <PrimaryButton
            onClick={handleGenerate}
            disabled={!keywords.trim() || isDrafting}
            onRenderIcon={() =>
              isDrafting ? <Spinner size={SpinnerSize.xSmall} /> : null
            }
          >
            {isDrafting
              ? strings.CopilotGeneratingLabel
              : strings.CopilotGenerateButton}
          </PrimaryButton>
        </div>
          </div>
        </Callout>
      )}
    </>
  );
};
