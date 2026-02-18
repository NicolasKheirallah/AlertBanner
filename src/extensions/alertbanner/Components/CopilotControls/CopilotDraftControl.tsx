import * as React from "react";
import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Label,
  Dialog,
  DialogType,
  DialogFooter,
  ActionButton,
  IconButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import {
  SparkleRegular,
  Checkmark24Regular,
  Dismiss24Regular,
  ArrowSync24Regular,
  TextBulletListSquare24Regular,
  TextExpand24Regular,
  TextCollapse24Regular,
  Lightbulb24Regular,
  Keyboard24Regular,
  Info24Regular,
  Building24Regular,
  Alert24Regular,
  Chat24Regular,
  CheckmarkCircle24Filled,
} from "@fluentui/react-icons";
import { CopilotService, CopilotTone } from "../Services/CopilotService";
import { logger } from "../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import styles from "./CopilotDraftControl.module.scss";

export interface ICopilotDraftControlProps {
  copilotService: CopilotService;
  onDraftGenerated: (draft: string) => void;
  onError: (error: string) => void;
  disabled?: boolean;
  alertType?: string;
  priority?: string;
}

type DraftStage = "input" | "generating" | "preview";

const EXAMPLE_PROMPTS = [
  strings.CopilotExamplePrompt1,
  strings.CopilotExamplePrompt2,
  strings.CopilotExamplePrompt3,
  strings.CopilotExamplePrompt4,
];

export const CopilotDraftControl: React.FC<ICopilotDraftControlProps> = ({
  copilotService,
  onDraftGenerated,
  onError,
  disabled,
  alertType,
  priority,
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [stage, setStage] = React.useState<DraftStage>("input");
  const [keywords, setKeywords] = React.useState("");
  const [selectedExample, setSelectedExample] = React.useState<number | null>(null);
  const [tone, setTone] = React.useState<CopilotTone>("Professional");
  const [generatedDraft, setGeneratedDraft] = React.useState("");
  const [showTips, setShowTips] = React.useState(false);
  const [isRefining, setIsRefining] = React.useState(false);

  const charCount = keywords.length;
  const isOverLimit = charCount > 500;
  const draftCharCount = generatedDraft.length;
  const draftWordCount = generatedDraft.trim().split(/\s+/).filter(w => w.length > 0).length;

  const buildContextPrompt = (): string => {
    const contextParts: string[] = [];
    if (alertType && alertType !== "Info") {
      contextParts.push(`This is a "${alertType}" type alert.`);
    }
    if (priority && priority !== "Medium") {
      contextParts.push(`Priority level: ${priority}.`);
    }
    return contextParts.join(" ");
  };

  const handleGenerate = async (refinement?: string): Promise<void> => {
    const promptText = refinement || keywords.trim();
    if (!promptText) return;

    if (refinement) {
      setIsRefining(true);
    } else {
      setStage("generating");
    }
    
    try {
      const context = buildContextPrompt();
      
      let fullPrompt: string;
      if (refinement) {
        fullPrompt = `Instruction: ${promptText}\n\nCurrent draft to improve: "${generatedDraft}"`;
      } else {
        fullPrompt = `${context}\n\n${promptText}`;
      }
      
      const response = await copilotService.generateDraftWithContext(
        fullPrompt,
        tone,
        !!refinement
      );

      if (response.isError) {
        if (!response.isCancelled) {
          onError(response.errorMessage || strings.CopilotDraftGenerationFailed);
        }
        if (!refinement) {
          setStage("input");
        }
      } else {
        setGeneratedDraft(response.content);
        setStage("preview");
        setIsRefining(false);
      }
    } catch (error) {
      logger.error("CopilotDraftControl", "Draft generation failed", error);
      onError(strings.CopilotUnexpectedError);
      if (!refinement) {
        setStage("input");
      }
      setIsRefining(false);
    }
  };

  const handleAccept = (): void => {
    onDraftGenerated(generatedDraft);
    setIsOpen(false);
    resetState();
  };

  const handleEdit = (): void => {
    setStage("input");
    setGeneratedDraft("");
    setIsRefining(false);
  };

  const handleRefine = (instruction: string): void => {
    void handleGenerate(instruction);
  };

  const handleExampleClick = (example: string, index: number): void => {
    setKeywords(example);
    setSelectedExample(index);
  };

  const resetState = (): void => {
    setStage("input");
    setKeywords("");
    setSelectedExample(null);
    setGeneratedDraft("");
    setShowTips(false);
    setIsRefining(false);
  };

  const handleCancel = (): void => {
    copilotService.cancelActiveOperation();
    setIsOpen(false);
    resetState();
  };

  const handleOpen = (): void => {
    if (!isOpen) {
      resetState();
    }
    setIsOpen(true);
  };

  // Keyboard shortcut: Ctrl+Enter to generate
  const handleKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      if (keywords.trim() && stage === "input") {
        void handleGenerate();
      }
    }
  };

  // Typing effect for generated text
  const [displayedText, setDisplayedText] = React.useState("");
  const [showCursor, setShowCursor] = React.useState(true);
  const typingRef = React.useRef<number | null>(null);
  const isTypingRef = React.useRef(false);
  
  React.useEffect(() => {
    // Cleanup function
    return () => {
      if (typingRef.current) {
        cancelAnimationFrame(typingRef.current);
      }
      isTypingRef.current = false;
    };
  }, []);
  
  React.useEffect(() => {
    if (stage === "preview" && generatedDraft) {
      // Cancel any existing animation
      if (typingRef.current) {
        cancelAnimationFrame(typingRef.current);
      }
      
      let index = 0;
      let lastTime = 0;
      const charDelay = 12;
      
      setDisplayedText("");
      setShowCursor(true);
      isTypingRef.current = true;
      
      const typeChar = (currentTime: number) => {
        if (!isTypingRef.current) return;
        
        if (currentTime - lastTime >= charDelay) {
          if (index <= generatedDraft.length) {
            setDisplayedText(generatedDraft.slice(0, index));
            index++;
            lastTime = currentTime;
          }
        }
        
        if (index <= generatedDraft.length) {
          typingRef.current = requestAnimationFrame(typeChar);
        } else {
          // Typing complete - hide cursor after a delay
          setTimeout(() => {
            if (isTypingRef.current) {
              setShowCursor(false);
            }
          }, 500);
        }
      };
      
      typingRef.current = requestAnimationFrame(typeChar);
      
      return () => {
        isTypingRef.current = false;
        if (typingRef.current) {
          cancelAnimationFrame(typingRef.current);
        }
      };
    }
  }, [stage, generatedDraft]);

  // Handle keyboard navigation for examples
  const handleExampleKeyDown = (e: React.KeyboardEvent, example: string, index: number): void => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      handleExampleClick(example, index);
    }
  };

  return (
    <>
      <DefaultButton
        onRenderIcon={() => <SparkleRegular className={styles.sparkleIcon} />}
        disabled={disabled || stage === "generating"}
        onClick={handleOpen}
        className={styles.ghostButton}
      >
        {strings.CopilotDraftButton}
      </DefaultButton>

      <Dialog
        hidden={!isOpen}
        onDismiss={handleCancel}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: (
            <div className={styles.dialogHeader}>
              <SparkleRegular className={styles.headerIcon} />
              <span>{strings.CopilotDraftTitle}</span>
            </div>
          ) as unknown as string,
        }}
        modalProps={{
          isBlocking: stage === "generating",
          className: `${styles.dialog} CopilotDraftDialog`,
        }}
        minWidth={640}
        maxWidth={720}
      >
        <div className={styles.dialogContent}>
          {stage === "input" && (
            <>
              {/* Tips Toggle */}
              <button 
                className={styles.tipsToggle}
                onClick={() => setShowTips(!showTips)}
                type="button"
                aria-expanded={showTips}
              >
                <Info24Regular className={styles.tipsIcon} />
                <span>{strings.CopilotTipsToggleLabel}</span>
                {showTips && (
                  <button 
                    className={styles.tipsClose}
                    onClick={(e) => {
                      e.stopPropagation();
                      setShowTips(false);
                    }}
                    type="button"
                    aria-label={strings.CopilotTipsCloseAriaLabel}
                  >
                    <Dismiss24Regular />
                  </button>
                )}
              </button>
              
              {showTips && (
                <div className={styles.tipsPanel}>
                  <ul>
                    <li>{strings.CopilotTipsItem1}</li>
                    <li>{strings.CopilotTipsItem2}</li>
                    <li>{strings.CopilotTipsItem3}</li>
                    <li>{strings.CopilotTipsItem4}</li>
                  </ul>
                </div>
              )}

              {/* Example Prompts */}
              <div className={styles.examplesSection}>
                <Label className={styles.sectionLabel}>{strings.CopilotExamplesLabel}</Label>
                <div className={styles.exampleChips} role="listbox" aria-label="Example prompts">
                  {EXAMPLE_PROMPTS.map((example, idx) => (
                    <button
                      key={idx}
                      className={`${styles.exampleChip} ${selectedExample === idx ? styles.exampleChipSelected : ""}`}
                      onClick={() => handleExampleClick(example, idx)}
                      onKeyDown={(e) => handleExampleKeyDown(e, example, idx)}
                      type="button"
                      role="option"
                      aria-selected={selectedExample === idx}
                      tabIndex={0}
                    >
                      {selectedExample === idx && (
                        <CheckmarkCircle24Filled className={styles.chipCheckmark} />
                      )}
                      <Lightbulb24Regular className={styles.chipIcon} />
                      {example}
                    </button>
                  ))}
                </div>
              </div>

              {/* Input Area */}
              <div className={styles.fieldGroup}>
                <div className={styles.labelRow}>
                  <Label htmlFor="copilot-keywords" className={styles.inputLabel}>
                    {strings.CopilotInputLabel}
                  </Label>
                  <span className={`${styles.charCount} ${isOverLimit ? styles.charCountOver : ""}`}>
                    {charCount}/500
                  </span>
                </div>
                <TextField
                  id="copilot-keywords"
                  placeholder={strings.CopilotInputPlaceholder}
                  value={keywords}
                  multiline
                  rows={4}
                  onChange={(_, value) => {
                    setKeywords(value || "");
                    setSelectedExample(null);
                  }}
                  onKeyDown={handleKeyDown}
                  className={styles.fullWidth}
                  errorMessage={isOverLimit ? strings.CopilotCharLimitError : undefined}
                />
                <div className={styles.keyboardHint}>
                  <Keyboard24Regular className={styles.keyboardIcon} />
                  <span>{strings.CopilotKeyboardHint}</span>
                </div>
              </div>

              {/* Tone Selector */}
              <div className={styles.toneSection}>
                <Label className={styles.sectionLabel}>{strings.CopilotToneSelectorLabel}</Label>
                <div className={styles.tonePills} role="radiogroup" aria-label="Select tone">
                  <button
                    type="button"
                    className={`${styles.tonePill} ${tone === "Professional" ? styles.tonePillActive : ""}`}
                    onClick={() => setTone("Professional")}
                    role="radio"
                    aria-checked={tone === "Professional"}
                  >
                    <Building24Regular className={styles.toneIcon} />
                    <span className={styles.toneLabel}>{strings.CopilotToneProfessionalShort}</span>
                  </button>
                  <button
                    type="button"
                    className={`${styles.tonePill} ${tone === "Urgent" ? styles.tonePillActive : ""}`}
                    onClick={() => setTone("Urgent")}
                    role="radio"
                    aria-checked={tone === "Urgent"}
                  >
                    <Alert24Regular className={styles.toneIcon} />
                    <span className={styles.toneLabel}>{strings.CopilotToneUrgentShort}</span>
                  </button>
                  <button
                    type="button"
                    className={`${styles.tonePill} ${tone === "Casual" ? styles.tonePillActive : ""}`}
                    onClick={() => setTone("Casual")}
                    role="radio"
                    aria-checked={tone === "Casual"}
                  >
                    <Chat24Regular className={styles.toneIcon} />
                    <span className={styles.toneLabel}>{strings.CopilotToneCasualShort}</span>
                  </button>
                </div>
              </div>
            </>
          )}

          {stage === "generating" && (
            <div className={styles.generatingState}>
              <div className={styles.generatingAnimation}>
                <div className={styles.dot}></div>
                <div className={styles.dot}></div>
                <div className={styles.dot}></div>
              </div>
              <p className={styles.generatingText}>
                {strings.CopilotGeneratingLabel}
              </p>
              <p className={styles.generatingSubtext}>
                {Text.format(strings.CopilotGeneratingSubtext, tone.toLowerCase())}
              </p>
            </div>
          )}

          {stage === "preview" && (
            <>
              <div className={styles.previewHeader}>
                <div className={styles.previewTitleGroup}>
                  <h3 className={styles.previewTitle}>{strings.CopilotPreviewTitle}</h3>
                  <span className={`${styles.previewBadge} ${styles[`previewBadge${tone}`]}`}>
                    {Text.format(strings.CopilotPreviewBadge, tone)}
                  </span>
                </div>
              </div>

              {/* Alert-style Preview */}
              <div className={styles.previewContainer}>
                <div className={`${styles.alertPreview} ${styles[`alertPreview${tone}`]}`}>
                  <div className={styles.alertIcon}>
                    {tone === "Professional" && <Building24Regular />}
                    {tone === "Urgent" && <Alert24Regular />}
                    {tone === "Casual" && <Chat24Regular />}
                  </div>
                  <div className={styles.alertContent}>
                    <p className={styles.previewText}>
                      {displayedText}
                      {showCursor && <span className={styles.cursor}>|</span>}
                    </p>
                  </div>
                </div>
              </div>

              {/* Draft Stats */}
              <div className={styles.draftStats}>
                <span className={styles.draftStat}>{Text.format(strings.CopilotDraftStatsWords, draftWordCount)}</span>
                <span className={styles.draftStatDivider}>|</span>
                <span className={styles.draftStat}>{Text.format(strings.CopilotDraftStatsChars, draftCharCount)}</span>
              </div>

              {/* Refine Actions */}
              <div className={styles.refineSection}>
                <Label className={styles.sectionLabel}>{strings.CopilotRefineSectionLabel}</Label>
                <div className={styles.refineActions}>
                  <ActionButton
                    iconProps={{ iconName: undefined }}
                    onRenderIcon={() => isRefining ? <Spinner size={SpinnerSize.xSmall} /> : <TextCollapse24Regular />}
                    onClick={() => handleRefine("Make this shorter and more concise. Keep only the essential information.")}
                    className={styles.refineButton}
                    disabled={isRefining}
                  >
                    {isRefining ? strings.CopilotRefiningLabel : strings.CopilotRefineShorter}
                  </ActionButton>
                  <ActionButton
                    iconProps={{ iconName: undefined }}
                    onRenderIcon={() => isRefining ? <Spinner size={SpinnerSize.xSmall} /> : <TextExpand24Regular />}
                    onClick={() => handleRefine("Make this longer with more details and context.")}
                    className={styles.refineButton}
                    disabled={isRefining}
                  >
                    {isRefining ? strings.CopilotRefiningLabel : strings.CopilotRefineLonger}
                  </ActionButton>
                  <ActionButton
                    iconProps={{ iconName: undefined }}
                    onRenderIcon={() => isRefining ? <Spinner size={SpinnerSize.xSmall} /> : <TextBulletListSquare24Regular />}
                    onClick={() => handleRefine("Rephrase this to be more engaging and impactful.")}
                    className={styles.refineButton}
                    disabled={isRefining}
                  >
                    {isRefining ? strings.CopilotRefiningLabel : strings.CopilotRefineRephrase}
                  </ActionButton>
                  <ActionButton
                    iconProps={{ iconName: undefined }}
                    onRenderIcon={() => isRefining ? <Spinner size={SpinnerSize.xSmall} /> : <ArrowSync24Regular />}
                    onClick={() => handleRefine("Generate a different variation on the same topic. Keep the same key message but use different wording.")}
                    className={styles.refineButton}
                    disabled={isRefining}
                  >
                    {isRefining ? strings.CopilotRefiningLabel : strings.CopilotRefineTryAgain}
                  </ActionButton>
                </div>
              </div>
            </>
          )}
        </div>

        <DialogFooter>
          {stage === "input" && (
            <>
              <DefaultButton onClick={handleCancel} text={strings.Cancel} />
              <PrimaryButton
                onClick={() => void handleGenerate()}
                disabled={!keywords.trim() || isOverLimit}
                iconProps={{ iconName: undefined }}
                onRenderIcon={() => <SparkleRegular className={styles.buttonSparkle} />}
              >
                {strings.CopilotGenerateButton}
              </PrimaryButton>
            </>
          )}
          
          {stage === "preview" && (
            <>
              <DefaultButton
                onClick={handleEdit}
                iconProps={{ iconName: undefined }}
                onRenderIcon={() => <Dismiss24Regular />}
              >
                {strings.CopilotEditButton}
              </DefaultButton>
              <PrimaryButton
                onClick={handleAccept}
                iconProps={{ iconName: undefined }}
                onRenderIcon={() => <Checkmark24Regular />}
              >
                {strings.CopilotAcceptDraft}
              </PrimaryButton>
            </>
          )}
        </DialogFooter>
      </Dialog>
    </>
  );
};
