import * as React from "react";
import { DefaultButton, IconButton, ActionButton, Spinner, SpinnerSize } from "@fluentui/react";
import {
  AppsListDetail24Regular,
  CheckmarkCircle24Filled,
  Warning24Filled,
  ErrorCircle24Filled,
  Dismiss24Regular,
  Lightbulb24Regular,
  Clock24Regular,
  TextWordCount24Regular,
  Building24Regular,
  Alert24Regular,
  Chat24Regular,
  Wand24Regular,
  ChevronDown24Regular,
  ChevronUp24Regular,
} from "@fluentui/react-icons";
import { CopilotService, ISentimentResult } from "../Services/CopilotService";
import { logger } from "../Services/LoggerService";
import * as strings from "AlertBannerApplicationCustomizerStrings";
import { Text } from "@microsoft/sp-core-library";
import styles from "./CopilotSentimentControl.module.scss";

export interface ICopilotSentimentControlProps {
  copilotService: CopilotService;
  textToAnalyze: string;
  onError: (error: string) => void;
  disabled?: boolean;
  onApplyFix?: (fixedText: string) => void;
}

interface IReadingStats {
  wordCount: number;
  readingTimeSeconds: number;
}

interface IFixableIssue {
  id: string;
  type: "critical" | "recommended" | "optional";
  message: string;
  fixAction?: string;
  fixHandler?: () => string;
}

interface IDetectedTone {
  tone: 'professional' | 'urgent' | 'casual';
  label: string;
  icon: React.ReactNode;
  color: string;
}

// Calculate reading stats
const calculateReadingStats = (text: string): IReadingStats => {
  const words = text.trim().split(/\s+/).filter(w => w.length > 0);
  const wordCount = words.length;
  const avgReadingSpeed = 200;
  const readingTimeSeconds = Math.ceil((wordCount / avgReadingSpeed) * 60);
  
  return { wordCount, readingTimeSeconds };
};

// Detect tone from text
const detectTone = (text: string): IDetectedTone => {
  const lowerText = text.toLowerCase();
  
  const urgentWords = ['urgent', 'immediately', 'asap', 'critical', 'emergency', 'alert', 'warning', 'important', 'attention'];
  const hasUrgentWords = urgentWords.some(word => lowerText.includes(word));
  
  const casualWords = ['hey', 'hi', 'hello', 'thanks', 'cheers', 'btw', 'just', 'maybe', 'probably'];
  const hasCasualWords = casualWords.some(word => lowerText.includes(word));
  
  const formalWords = ['please', 'regards', 'sincerely', 'respectfully', 'announcement', 'policy', 'procedure'];
  const hasFormalWords = formalWords.some(word => lowerText.includes(word));
  
  if (hasUrgentWords) {
    return { 
      tone: 'urgent', 
      label: 'Urgent',
      icon: <Alert24Regular />,
      color: '#d13438'
    };
  }
  
  if (hasCasualWords && !hasFormalWords) {
    return { 
      tone: 'casual', 
      label: 'Casual',
      icon: <Chat24Regular />,
      color: '#107c10'
    };
  }
  
  return { 
    tone: 'professional', 
    label: 'Professional',
    icon: <Building24Regular />,
    color: '#0078d4'
  };
};

// Calculate actual score based on issues
const calculateScore = (
  result: ISentimentResult, 
  stats: IReadingStats,
  fixedCount: number,
  totalIssues: number
): number => {
  let score = 100;
  
  // Deduct for word count
  if (stats.wordCount > 100) score -= 15;
  else if (stats.wordCount > 50) score -= 5;
  
  // Deduct for reading time
  if (stats.readingTimeSeconds > 20) score -= 10;
  else if (stats.readingTimeSeconds > 15) score -= 5;
  
  // Deduct for issues
  const criticalIssues = result.issues.filter(i => 
    i.toLowerCase().includes('missing') || 
    i.toLowerCase().includes('required')
  ).length;
  
  const recommendedIssues = result.issues.filter(i => 
    !i.toLowerCase().includes('missing') && 
    !i.toLowerCase().includes('required')
  ).length;
  
  score -= criticalIssues * 15;
  score -= recommendedIssues * 5;
  
  // Add back for fixed issues
  if (totalIssues > 0) {
    score += Math.round((fixedCount / totalIssues) * 20);
  }
  
  // Deduct for non-green status
  if (result.status === 'red') score -= 15;
  else if (result.status === 'yellow') score -= 5;
  
  return Math.max(0, Math.min(100, score));
};

// Generate fixable issues
const generateFixableIssues = (
  result: ISentimentResult, 
  stats: IReadingStats,
  text: string
): IFixableIssue[] => {
  const issues: IFixableIssue[] = [];
  
  if (stats.wordCount > 100) {
    issues.push({
      id: "too-long",
      type: "critical",
      message: `Too long for banner (${stats.wordCount} words). Aim for under 50.`,
      fixAction: strings.CopilotSentimentFixTrim,
      fixHandler: () => text.split(/[.!?]+/).slice(0, 2).join(". ") + ".",
    });
  }
  
  if (stats.readingTimeSeconds > 20) {
    issues.push({
      id: "too-slow",
      type: "recommended",
      message: `Takes ${stats.readingTimeSeconds}s to read. Consider shortening.`,
      fixAction: strings.CopilotSentimentFixConcise,
    });
  }
  
  result.issues.forEach((issue, idx) => {
    const lowerIssue = issue.toLowerCase();
    let type: "critical" | "recommended" | "optional" = "recommended";
    
    if (lowerIssue.includes("missing") || lowerIssue.includes("required") || lowerIssue.includes("no ")) {
      type = "critical";
    } else if (lowerIssue.includes("consider") || lowerIssue.includes("could") || lowerIssue.includes("might")) {
      type = "optional";
    }
    
    issues.push({
      id: `analysis-${idx}`,
      type,
      message: issue,
    });
  });
  
  return issues;
};

// Parse analysis content
const parseAnalysisContent = (rawContent: string) => {
  const lines = rawContent.split("\n").map((l) => l.trim());
  const sections: { 
    professional?: string; 
    tone?: string; 
    issues?: string; 
    status?: string; 
    summary?: string;
  } = {};
  
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (lower.startsWith("professional:")) {
      sections.professional = line.substring("professional:".length).trim();
    } else if (lower.startsWith("tone_appropriate:")) {
      sections.tone = line.substring("tone_appropriate:".length).trim();
    } else if (lower.startsWith("issues:")) {
      sections.issues = line.substring("issues:".length).trim();
    } else if (lower.startsWith("status:")) {
      sections.status = line.substring("status:".length).trim();
    } else if (lower.startsWith("summary:")) {
      sections.summary = line.substring("summary:".length).trim();
    }
  }
  
  return sections;
};

export const CopilotSentimentControl: React.FC<ICopilotSentimentControlProps> = ({ 
  copilotService, 
  textToAnalyze, 
  onError, 
  disabled,
  onApplyFix,
}) => {
  const [isChecking, setIsChecking] = React.useState(false);
  const [result, setResult] = React.useState<ISentimentResult | null>(null);
  const [isExpanded, setIsExpanded] = React.useState(false);
  const [appliedFixes, setAppliedFixes] = React.useState<Set<string>>(new Set());
  const [calculatedScore, setCalculatedScore] = React.useState<number>(0);
  const lastCheckedRef = React.useRef<string>("");

  const readingStats = React.useMemo(() => {
    if (!textToAnalyze) return null;
    return calculateReadingStats(textToAnalyze);
  }, [textToAnalyze]);
  
  const detectedTone = React.useMemo(() => {
    if (!textToAnalyze) return null;
    return detectTone(textToAnalyze);
  }, [textToAnalyze]);

  const performCheck = React.useCallback(async (force: boolean = false): Promise<void> => {
    if (!textToAnalyze || textToAnalyze.trim().length < 10) {
      setResult(null);
      return;
    }

    if (!force && textToAnalyze === lastCheckedRef.current) return;
    lastCheckedRef.current = textToAnalyze;

    setIsChecking(true);
    try {
      const response = await copilotService.analyzeSentiment(textToAnalyze);

      if (response.isError) {
        if (!response.isCancelled) {
          onError(response.errorMessage || strings.CopilotAnalyzeFailed);
        }
        lastCheckedRef.current = "";
        setResult(null);
      } else {
        const parsed = copilotService.parseSentimentResult(response.content);
        setResult(parsed);
        setAppliedFixes(new Set());
        setIsExpanded(parsed.status !== "green" || parsed.issues.length > 0);
      }
    } catch (error) {
      logger.error("CopilotSentimentControl", "Sentiment check failed", error);
      onError(strings.CopilotUnexpectedError);
      lastCheckedRef.current = "";
    } finally {
      setIsChecking(false);
    }
  }, [textToAnalyze, copilotService, onError]);

  const handleCheck = (): void => {
    void performCheck(true);
  };

  const handleDismiss = (): void => {
    setResult(null);
    lastCheckedRef.current = "";
  };

  const handleApplyFix = (issue: IFixableIssue): void => {
    if (issue.fixHandler && onApplyFix) {
      const fixedText = issue.fixHandler();
      onApplyFix(fixedText);
      setAppliedFixes(prev => new Set(prev).add(issue.id));
      
      // Auto re-check after fix is applied
      setTimeout(() => {
        lastCheckedRef.current = ""; // Force re-check
        void performCheck(true);
      }, 500);
    }
  };

  const analysisSections = result ? parseAnalysisContent(result.rawContent) : null;
  const fixableIssues = result && readingStats 
    ? generateFixableIssues(result, readingStats, textToAnalyze)
    : [];
  
  const criticalIssues = fixableIssues.filter(i => i.type === "critical" && !appliedFixes.has(i.id));
  const recommendedIssues = fixableIssues.filter(i => i.type === "recommended" && !appliedFixes.has(i.id));
  const optionalIssues = fixableIssues.filter(i => i.type === "optional" && !appliedFixes.has(i.id));
  
  const totalIssues = fixableIssues.length;
  const fixedCount = appliedFixes.size;
  
  // Calculate actual score
  const score = result && readingStats 
    ? calculateScore(result, readingStats, fixedCount, totalIssues)
    : 0;
  
  // Determine status based on score
  const getStatusFromScore = (s: number): 'green' | 'yellow' | 'red' => {
    if (s >= 80) return 'green';
    if (s >= 50) return 'yellow';
    return 'red';
  };
  
  const status = getStatusFromScore(score);
  
  // Status config based on calculated score
  const statusConfig = {
    green: {
      icon: <CheckmarkCircle24Filled />,
      className: styles.statusGreen,
      barClassName: styles.barGreen,
      label: 'Looks Good',
      color: '#107c10',
    },
    yellow: {
      icon: <Warning24Filled />,
      className: styles.statusYellow,
      barClassName: styles.barYellow,
      label: 'Review Suggested',
      color: '#f9a825',
    },
    red: {
      icon: <ErrorCircle24Filled />,
      className: styles.statusRed,
      barClassName: styles.barRed,
      label: 'Issues Found',
      color: '#d13438',
    },
  };
  
  const config = statusConfig[status];
  const progressPercent = totalIssues > 0 ? Math.round((fixedCount / totalIssues) * 100) : 100;

  // Compact pill view when minimized
  const showPill = result && !isExpanded;

  if (!textToAnalyze || textToAnalyze.trim().length < 10) {
    return (
      <div className={styles.container}>
        <DefaultButton
          onRenderIcon={() => <AppsListDetail24Regular />}
          onClick={handleCheck}
          disabled={true}
          className={styles.sentimentButton}
          title={strings.CopilotSentimentTooltipDisabled}
        >
          {strings.CopilotSentimentButton}
        </DefaultButton>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.buttonRow}>
        <DefaultButton
          onRenderIcon={() => isChecking ? <Spinner size={SpinnerSize.xSmall} /> : <AppsListDetail24Regular />}
          onClick={handleCheck}
          disabled={disabled || isChecking}
          className={styles.sentimentButton}
        >
          {isChecking ? 'Checking...' : strings.CopilotSentimentButton}
        </DefaultButton>
        
        {showPill && (
          <button 
            className={`${styles.scorePill} ${config.className}`}
            onClick={() => setIsExpanded(true)}
            title={strings.CopilotSentimentTooltipExpand}
            type="button"
          >
            <span className={styles.pillIcon}>{config.icon}</span>
            <span className={styles.pillScore}>{score}%</span>
            {fixedCount > 0 && (
              <span className={styles.pillProgress}> ({fixedCount}/{totalIssues} fixed)</span>
            )}
            <ChevronDown24Regular className={styles.pillChevron} />
          </button>
        )}
      </div>

      {result && config && analysisSections && isExpanded && (
        <div className={`${styles.resultCard} ${config.className}`}>
          {/* Header */}
          <div className={styles.resultHeader}>
            <div className={styles.statusBadge}>
              <span className={styles.statusIcon}>{config.icon}</span>
              <div className={styles.statusInfo}>
                <span className={styles.statusLabel}>
                  {config.label} • {score}/100
                </span>
                <div className={styles.scoreBar}>
                  <div 
                    className={`${styles.scoreBarFill} ${config.barClassName}`}
                    style={{ width: `${score}%` }}
                  />
                </div>
              </div>
            </div>
            <div className={styles.headerActions}>
              <button
                className={styles.collapseButton}
                onClick={() => setIsExpanded(false)}
                title={strings.CopilotSentimentTooltipMinimize}
                type="button"
              >
                <ChevronUp24Regular />
              </button>
              <IconButton
                iconProps={{ iconName: undefined }}
                onClick={handleDismiss}
                className={styles.dismissButton}
                title={strings.CopilotSentimentTooltipDismiss}
                ariaLabel={strings.CopilotSentimentDismissAriaLabel}
                onRenderIcon={() => <Dismiss24Regular />}
              />
            </div>
          </div>

          {/* Reading Stats */}
          {readingStats && detectedTone && (
            <div className={styles.statsBar}>
              <div className={styles.statItem}>
                <TextWordCount24Regular className={styles.statIcon} />
                <span className={styles.statValue}>{readingStats.wordCount}</span>
                <span className={styles.statLabel}>{strings.CopilotSentimentReadingStatsWords}</span>
              </div>
              <div className={styles.statDivider} />
              <div className={styles.statItem}>
                <Clock24Regular className={styles.statIcon} />
                <span className={styles.statValue}>{readingStats.readingTimeSeconds}s</span>
                <span className={styles.statLabel}>{strings.CopilotSentimentReadingStatsTime}</span>
              </div>
              <div className={styles.statDivider} />
              <div className={styles.statItem}>
                <span className={styles.toneIcon} style={{ color: detectedTone.color }}>
                  {detectedTone.icon}
                </span>
                <span className={styles.statValue}>{detectedTone.label}</span>
                <span className={styles.statLabel}>{strings.CopilotSentimentToneLabel}</span>
              </div>
            </div>
          )}

          {/* Progress Checklist */}
          {totalIssues > 0 && (
            <div className={styles.progressSection}>
              <div className={styles.progressHeader}>
                <span className={styles.progressLabel}>{strings.CopilotSentimentIssuesFixedLabel}</span>
                <span className={styles.progressValue}>{Text.format(strings.CopilotSentimentIssuesFixedCount, fixedCount, totalIssues)}</span>
              </div>
              <div className={styles.progressBar}>
                <div 
                  className={styles.progressFill}
                  style={{ width: `${progressPercent}%` }}
                />
              </div>
            </div>
          )}

          {/* Critical Issues */}
          {criticalIssues.length > 0 && (
            <div className={styles.issuesSectionCritical}>
              <h4 className={styles.sectionTitleCritical}>
                <ErrorCircle24Filled className={styles.sectionIcon} />
                {strings.CopilotSentimentCriticalSection} ({criticalIssues.length})
              </h4>
              <ul className={styles.issueList}>
                {criticalIssues.map((issue) => (
                  <li key={issue.id} className={styles.issueItem}>
                    <div className={styles.issueContent}>
                      <span className={styles.issueText}>{issue.message}</span>
                      {issue.fixAction && onApplyFix && (
                        <ActionButton
                          className={styles.fixButton}
                          onClick={() => handleApplyFix(issue)}
                          iconProps={{ iconName: undefined }}
                          onRenderIcon={() => <Wand24Regular className={styles.fixIcon} />}
                        >
                          {issue.fixAction}
                        </ActionButton>
                      )}
                    </div>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Recommended Issues */}
          {recommendedIssues.length > 0 && (
            <div className={styles.issuesSectionRecommended}>
              <h4 className={styles.sectionTitleRecommended}>
                <Warning24Filled className={styles.sectionIcon} />
                {strings.CopilotSentimentRecommendedSection} ({recommendedIssues.length})
              </h4>
              <ul className={styles.issueList}>
                {recommendedIssues.map((issue) => (
                  <li key={issue.id} className={styles.issueItem}>
                    <div className={styles.issueContent}>
                      <span className={styles.issueText}>{issue.message}</span>
                      {issue.fixAction && onApplyFix && (
                        <ActionButton
                          className={styles.fixButton}
                          onClick={() => handleApplyFix(issue)}
                          iconProps={{ iconName: undefined }}
                          onRenderIcon={() => <Wand24Regular className={styles.fixIcon} />}
                        >
                          {issue.fixAction}
                        </ActionButton>
                      )}
                    </div>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Optional Issues */}
          {optionalIssues.length > 0 && (
            <div className={styles.issuesSectionOptional}>
              <h4 className={styles.sectionTitleOptional}>
                <Lightbulb24Regular className={styles.sectionIcon} />
                {strings.CopilotSentimentOptionalSection} ({optionalIssues.length})
              </h4>
              <ul className={styles.issueList}>
                {optionalIssues.map((issue) => (
                  <li key={issue.id} className={styles.issueItemOptional}>
                    {issue.message}
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Analysis Summary */}
          <div className={styles.analysisGrid}>
            {analysisSections.professional && (
              <div className={styles.analysisItem}>
                <span className={styles.analysisLabel}>{strings.CopilotSentimentAnalysisProfessional}</span>
                <span className={`${styles.analysisValue} ${
                  analysisSections.professional.toLowerCase().startsWith("yes") 
                    ? styles.valuePositive 
                    : styles.valueNegative
                }`}>
                  {analysisSections.professional}
                </span>
              </div>
            )}
            {analysisSections.tone && (
              <div className={styles.analysisItem}>
                <span className={styles.analysisLabel}>{strings.CopilotSentimentAnalysisTone}</span>
                <span className={`${styles.analysisValue} ${
                  analysisSections.tone.toLowerCase().startsWith("yes") 
                    ? styles.valuePositive 
                    : styles.valueNegative
                }`}>
                  {analysisSections.tone}
                </span>
              </div>
            )}
          </div>

          {/* Summary */}
          {analysisSections.summary && (
            <div className={styles.summarySection}>
              <p>{analysisSections.summary}</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
