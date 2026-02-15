import { MSGraphClientV3 } from "@microsoft/sp-http";
import { logger } from "./LoggerService";

export interface ICopilotResponse {
  content: string;
  isError: boolean;
  errorMessage?: string;
  citations?: ICopilotCitation[];
}

/**
 * Citation returned by Copilot when grounding with enterprise/web data.
 */
export interface ICopilotCitation {
  title: string;
  url?: string;
  snippet?: string;
}

/**
 * Supported tones for Copilot draft generation.
 */
export type CopilotTone = "Professional" | "Urgent" | "Casual";

/**
 * Structured governance analysis result parsed from Copilot response.
 */
export interface IGovernanceResult {
  isProfessional: boolean;
  isToneAppropriate: boolean;
  issues: string[];
  status: "green" | "yellow" | "red";
  rawContent: string;
}

/**
 * File context for grounding Copilot responses with SharePoint/OneDrive data.
 */
export interface ICopilotFileContext {
  id: string;
  source: "oneDrive" | "sharePoint";
}

/**
 * Metadata for a Copilot conversation returned by the API.
 */
interface IConversationMetadata {
  id: string;
  createdDateTime: string;
  status: string;
  turnCount: number;
}

/**
 * Service for interacting with the Microsoft 365 Copilot Chat API via Microsoft Graph.
 *
 * API Reference: https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-beta
 *
 * IMPORTANT: This uses the /beta/copilot/conversations endpoint which is
 * in preview and subject to change. Not recommended for production use.
 *
 * Endpoints:
 *   POST /beta/copilot/conversations        → Create a new conversation
 *   POST /beta/copilot/conversations/{id}/chat → Send a message
 *
 * Required delegated permissions:
 *   - Sites.ReadWrite.All (already granted for list operations)
 *   - User.Read
 *   - Mail.Send
 *   - GroupMember.Read.All
 *   - Directory.Read.All
 *
 * Optional Copilot grounding permissions (expand what data Copilot can access):
 *   - Mail.Read, People.Read.All, Chat.Read, ChannelMessage.Read.All,
 *     ExternalItem.Read.All — not required for alert drafting/translation
 *
 * Each user must have a Microsoft 365 Copilot license.
 *
 * @file CopilotService.ts
 * @author Nicolas Kheirallah
 * @version 6.0.0
 * @since 2026-02-15
 */
export class CopilotService {
  private graphClient: MSGraphClientV3;
  private readonly endpoint = "/copilot/conversations";
  private cachedConversation: IConversationMetadata | null = null;
  private activeAbortController: AbortController | null = null;

  /** Maximum conversation age (30 minutes) before forcing a new one. */
  private static readonly MAX_CONVERSATION_AGE_MS = 30 * 60 * 1000;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Sanitizes user input to prevent prompt injection attacks.
   * Escapes characters that could alter the AI prompt structure.
   *
   * @param input - Raw user input
   * @returns Sanitized string safe for prompt interpolation
   */
  private sanitizeInput(input: string): string {
    return input
      .replace(/\\/g, "\\\\")
      .replace(/"/g, '\\"')
      .replace(/\n/g, " ")
      .replace(/\r/g, "")
      .trim();
  }

  /**
   * Generates a draft alert based on the provided user prompt (keywords).
   *
   * @param keywords - The user's input keywords or rough draft
   * @param tone - The desired tone for the generated draft
   * @returns The AI-generated draft text
   */
  public async generateDraft(
    keywords: string,
    tone: CopilotTone = "Professional",
  ): Promise<ICopilotResponse> {
    try {
      const conversationId = await this.ensureConversation();
      const sanitizedKeywords = this.sanitizeInput(keywords);

      const prompt = `You are an assistant for a SharePoint Administrator. 
      Draft a ${tone.toLowerCase()} alert banner message based on these keywords: "${sanitizedKeywords}".
      The message should be suitable for a corporate intranet. 
      Return ONLY the message text, no pleasantries or introductions.`;

      return await this.sendMessage(conversationId, prompt);
    } catch (error) {
      logger.error("CopilotService", "Failed to generate draft", error);
      return {
        content: "",
        isError: true,
        errorMessage: this.getFriendlyErrorMessage(error),
      };
    }
  }

  /**
   * Analyzes the sentiment and tone of the provided text.
   * Returns a structured governance result with a green/yellow/red status.
   *
   * @param text - The alert text to analyze
   * @returns Analysis response
   */
  public async analyzeSentiment(text: string): Promise<ICopilotResponse> {
    try {
      const conversationId = await this.ensureConversation();
      const sanitizedText = this.sanitizeInput(text);

      const prompt = `Analyze the following text for a corporate alert banner. 
      Text: "${sanitizedText}"
      
      Respond in EXACTLY this format (each item on its own line):
      PROFESSIONAL: Yes or No
      TONE_APPROPRIATE: Yes or No
      ISSUES: comma-separated list of issues, or "None"
      STATUS: Green, Yellow, or Red
      SUMMARY: one sentence summary of your analysis`;

      return await this.sendMessage(conversationId, prompt);
    } catch (error) {
      logger.error("CopilotService", "Failed to analyze sentiment", error);
      return {
        content: "",
        isError: true,
        errorMessage: this.getFriendlyErrorMessage(error),
      };
    }
  }

  /**
   * Parses a raw governance analysis response into a structured result.
   *
   * @param rawContent - The raw string response from analyzeSentiment
   * @returns Parsed governance result with typed fields
   */
  public parseGovernanceResult(rawContent: string): IGovernanceResult {
    const lines = rawContent.split("\n").map((l) => l.trim());
    const result: IGovernanceResult = {
      isProfessional: true,
      isToneAppropriate: true,
      issues: [],
      status: "green",
      rawContent,
    };

    for (const line of lines) {
      const lower = line.toLowerCase();
      if (lower.startsWith("professional:")) {
        result.isProfessional = lower.includes("yes");
      } else if (lower.startsWith("tone_appropriate:")) {
        result.isToneAppropriate = lower.includes("yes");
      } else if (lower.startsWith("issues:")) {
        const issuesStr = line.substring("issues:".length).trim();
        if (issuesStr.toLowerCase() !== "none" && issuesStr.length > 0) {
          result.issues = issuesStr.split(",").map((i) => i.trim());
        }
      } else if (lower.startsWith("status:")) {
        const statusStr = line.substring("status:".length).trim().toLowerCase();
        if (
          statusStr === "red" ||
          statusStr === "yellow" ||
          statusStr === "green"
        ) {
          result.status = statusStr;
        }
      }
    }

    return result;
  }

  /**
   * Translates the provided text to the target language.
   *
   * @param text - The text to translate
   * @param targetLanguage - The target language (e.g., "French", "German")
   * @returns The translated text
   */
  public async translateText(
    text: string,
    targetLanguage: string,
  ): Promise<ICopilotResponse> {
    try {
      const conversationId = await this.ensureConversation();
      const sanitizedText = this.sanitizeInput(text);
      const sanitizedLang = this.sanitizeInput(targetLanguage);

      const prompt = `Translate the following corporate alert message to ${sanitizedLang}. 
      Maintain the professional tone and urgency.
      Text: "${sanitizedText}"
      Return ONLY the translated text.`;

      return await this.sendMessage(conversationId, prompt);
    } catch (error) {
      logger.error("CopilotService", "Failed to translate text", error);
      return {
        content: "",
        isError: true,
        errorMessage: this.getFriendlyErrorMessage(error),
      };
    }
  }

  /**
   * Sends a message with SharePoint/OneDrive file context for grounding.
   * Copilot will use the file contents to produce more relevant responses.
   *
   * @param prompt - The message text
   * @param files - Array of file references for context grounding
   * @returns Copilot's grounded response
   */
  public async sendWithContext(
    prompt: string,
    files: ICopilotFileContext[],
  ): Promise<ICopilotResponse> {
    try {
      const conversationId = await this.ensureConversation();
      const sanitizedPrompt = this.sanitizeInput(prompt);

      return await this.sendMessage(conversationId, sanitizedPrompt, {
        files,
      });
    } catch (error) {
      logger.error(
        "CopilotService",
        "Failed to send message with context",
        error,
      );
      return {
        content: "",
        isError: true,
        errorMessage: this.getFriendlyErrorMessage(error),
      };
    }
  }

  /**
   * Checks if the current user has access to Copilot APIs
   * by attempting to create a conversation. A 403 typically means
   * either missing permissions or no M365 Copilot license.
   */
  public async checkAccess(): Promise<boolean> {
    try {
      await this.ensureConversation();
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes("401") || message.includes("403")) {
        logger.warn(
          "CopilotService",
          "Copilot access denied — user may lack M365 Copilot license or " +
            "required delegated permissions (Sites.Read.All, Mail.Read, " +
            "People.Read.All, OnlineMeetingTranscript.Read.All, Chat.Read, " +
            "ChannelMessage.Read.All, ExternalItem.Read.All)",
          error,
        );
      } else {
        logger.warn("CopilotService", "Copilot access check failed", error);
      }
      return false;
    }
  }

  /**
   * Cancels any active Copilot operation.
   */
  public cancelActiveOperation(): void {
    if (this.activeAbortController) {
      this.activeAbortController.abort();
      this.activeAbortController = null;
    }
  }

  /**
   * Clears the cached conversation, forcing a new conversation
   * on the next API call.
   */
  public resetConversation(): void {
    this.cachedConversation = null;
  }

  /**
   * Ensures a conversation exists, re-using a cached conversation
   * when available and not expired. Creates a new conversation if needed.
   */
  private async ensureConversation(): Promise<string> {
    if (this.cachedConversation && !this.isConversationExpired()) {
      return this.cachedConversation.id;
    }

    try {
      const res = await this.graphClient.api(`/beta${this.endpoint}`).post({});

      this.cachedConversation = {
        id: res.id,
        createdDateTime: res.createdDateTime || new Date().toISOString(),
        status: res.status || "active",
        turnCount: res.turnCount || 0,
      };

      return res.id;
    } catch (error) {
      logger.error("CopilotService", "Failed to create conversation", error);
      this.cachedConversation = null;
      throw error;
    }
  }

  /**
   * Checks if the cached conversation has exceeded the max age.
   */
  private isConversationExpired(): boolean {
    if (!this.cachedConversation) return true;

    const created = new Date(this.cachedConversation.createdDateTime).getTime();
    const age = Date.now() - created;

    return age > CopilotService.MAX_CONVERSATION_AGE_MS;
  }

  /**
   * Sends a message to a Copilot conversation via the /chat endpoint
   * and returns the reply.
   *
   * API: POST /beta/copilot/conversations/{id}/chat
   * Body: { message: { text: "..." }, context?: { files: [...] } }
   *
   * @param conversationId - The conversation to send to
   * @param content - The message text
   * @param context - Optional file context for grounding
   * @param enableWebGrounding - Whether to enable web search grounding
   */
  private async sendMessage(
    conversationId: string,
    content: string,
    context?: { files: ICopilotFileContext[] },
    enableWebGrounding: boolean = false,
  ): Promise<ICopilotResponse> {
    this.activeAbortController = new AbortController();

    const requestBody: Record<string, unknown> = {
      message: {
        text: content,
      },
    };

    if (context?.files && context.files.length > 0) {
      requestBody.context = {
        files: context.files.map((f) => ({
          id: f.id,
          source: f.source,
        })),
      };
    }

    if (enableWebGrounding) {
      requestBody.enableWebGrounding = true;
    }

    try {
      const res = await this.graphClient
        .api(`/beta${this.endpoint}/${conversationId}/chat`)
        .post(requestBody);

      // Update turn count in cached metadata
      if (this.cachedConversation) {
        this.cachedConversation.turnCount++;
      }

      // Parse response — the API returns a copilotConversationMessage
      const responseText = this.extractResponseText(res);
      const citations = this.extractCitations(res);

      return {
        content: responseText,
        isError: false,
        citations: citations.length > 0 ? citations : undefined,
      };
    } catch (error) {
      // If the conversation expired or was invalid, reset cache and retry once
      const errorMessage = (error as Error)?.message || "";
      if (errorMessage.includes("404") || errorMessage.includes("NotFound")) {
        logger.warn("CopilotService", "Conversation expired, creating new one");
        this.cachedConversation = null;
        const newConversationId = await this.ensureConversation();

        const retryRes = await this.graphClient
          .api(`/beta${this.endpoint}/${newConversationId}/chat`)
          .post(requestBody);

        const responseText = this.extractResponseText(retryRes);
        const citations = this.extractCitations(retryRes);

        return {
          content: responseText,
          isError: false,
          citations: citations.length > 0 ? citations : undefined,
        };
      }

      logger.error("CopilotService", "Failed to send message", error);
      throw error;
    } finally {
      this.activeAbortController = null;
    }
  }

  /**
   * Extracts the response text from the API response object.
   * Handles multiple possible response shapes from the beta API.
   */
  private extractResponseText(response: Record<string, unknown>): string {
    // Shape 1: { message: { text: "..." } }
    const message = response.message as Record<string, unknown> | undefined;
    if (message?.text) {
      return String(message.text);
    }

    // Shape 2: { value: [{ content: "..." }] }
    const value = response.value as Array<Record<string, unknown>> | undefined;
    if (value && value.length > 0 && value[0].content) {
      return String(value[0].content);
    }

    // Shape 3: { text: "..." }
    if (response.text) {
      return String(response.text);
    }

    // Shape 4: { content: "..." }
    if (response.content) {
      return String(response.content);
    }

    logger.warn(
      "CopilotService",
      "Unexpected response shape from Copilot API",
      response,
    );
    return JSON.stringify(response);
  }

  /**
   * Extracts citations from the API response, if present.
   * Citations indicate which enterprise or web data sources
   * Copilot used to ground its response.
   */
  private extractCitations(
    response: Record<string, unknown>,
  ): ICopilotCitation[] {
    const citations = (response.citations || response.groundingCitations) as
      | Array<Record<string, unknown>>
      | undefined;
    if (!citations || !Array.isArray(citations)) {
      return [];
    }

    return citations.map((c) => ({
      title: String(c.title || c.displayName || ""),
      url: c.url ? String(c.url) : undefined,
      snippet: c.snippet ? String(c.snippet) : undefined,
    }));
  }

  /**
   * Maps HTTP error codes to user-friendly messages.
   */
  private getFriendlyErrorMessage(error: unknown): string {
    const message = error instanceof Error ? error.message : String(error);

    if (message.includes("401") || message.includes("403")) {
      return "You do not have permission or a valid Microsoft 365 Copilot license to use this feature.";
    }
    if (message.includes("429")) {
      return "Copilot is busy. Please try again later.";
    }
    if (message.includes("400")) {
      return "Invalid request to Copilot. Please try rephrasing your input.";
    }
    return "An unexpected error occurred while communicating with Copilot.";
  }
}
