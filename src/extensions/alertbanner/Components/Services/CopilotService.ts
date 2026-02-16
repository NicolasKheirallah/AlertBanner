import { MSGraphClientV3 } from "@microsoft/sp-http";
import { logger } from "./LoggerService";

export interface ICopilotResponse {
  content: string;
  isError: boolean;
  isCancelled?: boolean;
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
  private activeAbortControllers = new Set<AbortController>();
  private copilotAvailability: "unknown" | "available" | "unavailable" =
    "unknown";

  /** Maximum conversation age (30 minutes) before forcing a new one. */
  private static readonly MAX_CONVERSATION_AGE_MS = 30 * 60 * 1000;
  private static readonly COPILOT_UNAVAILABLE_MESSAGE =
    "Copilot API is not available in this environment.";

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Builds a Copilot Graph request against the beta API version.
   * MSGraphClientV3 defaults to v1.0, so beta must be selected explicitly.
   */
  private copilotApi(pathSuffix: string = "") {
    return this.graphClient
      .api(`${this.endpoint}${pathSuffix}`)
      .version("beta");
  }

  /**
   * Performs basic normalization/escaping before interpolating user input
   * into prompt templates.
   *
   * @param input - Raw user input
   * @returns Sanitized string for prompt interpolation
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
    const abortController = this.beginOperation();
    try {
      const conversationId = await this.ensureConversation(
        false,
        abortController.signal,
      );
      const sanitizedKeywords = this.sanitizeInput(keywords);

      const prompt = `You are an assistant for a SharePoint Administrator. 
      Draft a ${tone.toLowerCase()} alert banner message based on these keywords: "${sanitizedKeywords}".
      The message should be suitable for a corporate intranet. 
      Return ONLY the message text, no pleasantries or introductions.`;

      return await this.sendMessage(
        conversationId,
        prompt,
        undefined,
        false,
        abortController.signal,
      );
    } catch (error) {
      if (!this.isAbortError(error)) {
        logger.error("CopilotService", "Failed to generate draft", error);
      }
      return this.createErrorResponse(error);
    } finally {
      this.endOperation(abortController);
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
    const abortController = this.beginOperation();
    try {
      const conversationId = await this.ensureConversation(
        false,
        abortController.signal,
      );
      const sanitizedText = this.sanitizeInput(text);

      const prompt = `Analyze the following text for a corporate alert banner. 
      Text: "${sanitizedText}"
      
      Respond in EXACTLY this format (each item on its own line):
      PROFESSIONAL: Yes or No
      TONE_APPROPRIATE: Yes or No
      ISSUES: comma-separated list of issues, or "None"
      STATUS: Green, Yellow, or Red
      SUMMARY: one sentence summary of your analysis`;

      return await this.sendMessage(
        conversationId,
        prompt,
        undefined,
        false,
        abortController.signal,
      );
    } catch (error) {
      if (!this.isAbortError(error)) {
        logger.error("CopilotService", "Failed to analyze sentiment", error);
      }
      return this.createErrorResponse(error);
    } finally {
      this.endOperation(abortController);
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
      isProfessional: false,
      isToneAppropriate: false,
      issues: [],
      status: "yellow",
      rawContent,
    };
    let professionalParsed = false;
    let toneParsed = false;
    let statusParsed = false;

    for (const line of lines) {
      const lower = line.toLowerCase();
      if (lower.startsWith("professional:")) {
        const value = this.parseBooleanLine(line);
        if (value !== undefined) {
          result.isProfessional = value;
          professionalParsed = true;
        }
      } else if (lower.startsWith("tone_appropriate:")) {
        const value = this.parseBooleanLine(line);
        if (value !== undefined) {
          result.isToneAppropriate = value;
          toneParsed = true;
        }
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
          statusParsed = true;
        }
      }
    }

    if (!statusParsed) {
      if (!professionalParsed || !toneParsed) {
        result.status = "yellow";
      } else if (result.issues.length > 0) {
        result.status = result.issues.length > 2 ? "red" : "yellow";
      } else if (!result.isProfessional || !result.isToneAppropriate) {
        result.status = "yellow";
      } else {
        result.status = "green";
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
    const abortController = this.beginOperation();
    try {
      const conversationId = await this.ensureConversation(
        false,
        abortController.signal,
      );
      const sanitizedText = this.sanitizeInput(text);
      const sanitizedLang = this.sanitizeInput(targetLanguage);

      const prompt = `Translate the following corporate alert message to ${sanitizedLang}. 
      Maintain the professional tone and urgency.
      Text: "${sanitizedText}"
      Return ONLY the translated text.`;

      return await this.sendMessage(
        conversationId,
        prompt,
        undefined,
        false,
        abortController.signal,
      );
    } catch (error) {
      if (!this.isAbortError(error)) {
        logger.error("CopilotService", "Failed to translate text", error);
      }
      return this.createErrorResponse(error);
    } finally {
      this.endOperation(abortController);
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
    const abortController = this.beginOperation();
    try {
      const conversationId = await this.ensureConversation(
        false,
        abortController.signal,
      );
      const sanitizedPrompt = this.sanitizeInput(prompt);

      return await this.sendMessage(conversationId, sanitizedPrompt, {
        files,
      }, false, abortController.signal);
    } catch (error) {
      if (!this.isAbortError(error)) {
        logger.error(
          "CopilotService",
          "Failed to send message with context",
          error,
        );
      }
      return this.createErrorResponse(error);
    } finally {
      this.endOperation(abortController);
    }
  }

  /**
   * Checks if the current user has access to Copilot APIs
   * by attempting to create a conversation. A 403 typically means
   * either missing permissions or no M365 Copilot license.
   */
  public async checkAccess(): Promise<boolean> {
    const abortController = this.beginOperation();
    try {
      if (this.copilotAvailability === "unavailable") {
        return false;
      }
      // Force a fresh API call to avoid stale true from cached conversation.
      await this.ensureConversation(true, abortController.signal);
      return true;
    } catch (error) {
      const statusCode = this.getErrorStatusCode(error);
      if (statusCode === 401 || statusCode === 403) {
        logger.warn(
          "CopilotService",
          "Copilot access denied — user may lack M365 Copilot license or required delegated permissions",
          error,
        );
      } else {
        logger.warn("CopilotService", "Copilot access check failed", error);
      }
      return false;
    } finally {
      this.endOperation(abortController);
    }
  }

  /**
   * Cancels any active Copilot operation.
   */
  public cancelActiveOperation(): void {
    if (this.activeAbortControllers.size === 0) {
      return;
    }
    this.activeAbortControllers.forEach((controller) => controller.abort());
    this.activeAbortControllers.clear();
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
  private async ensureConversation(
    forceRefresh: boolean = false,
    signal?: AbortSignal,
  ): Promise<string> {
    if (this.copilotAvailability === "unavailable") {
      throw new Error(CopilotService.COPILOT_UNAVAILABLE_MESSAGE);
    }

    if (!forceRefresh && this.cachedConversation && !this.isConversationExpired()) {
      return this.cachedConversation.id;
    }

    if (forceRefresh) {
      this.cachedConversation = null;
    }

    try {
      const request = this.copilotApi();
      if (signal) {
        request.option("signal", signal);
      }
      const res = await request.post({});

      const conversationId =
        res && typeof res.id === "string" ? res.id.trim() : "";
      if (!conversationId) {
        throw new Error("Copilot conversation creation returned no conversation id.");
      }

      this.cachedConversation = {
        id: conversationId,
        createdDateTime: res.createdDateTime || new Date().toISOString(),
        status: res.status || "active",
        turnCount: res.turnCount || 0,
      };
      this.copilotAvailability = "available";

      return conversationId;
    } catch (error) {
      if (this.isCopilotEndpointUnavailable(error)) {
        this.copilotAvailability = "unavailable";
        logger.warn(
          "CopilotService",
          "Copilot endpoint is unavailable for this tenant/environment",
          error,
        );
        throw new Error(CopilotService.COPILOT_UNAVAILABLE_MESSAGE);
      }
      if (!this.isAbortError(error)) {
        logger.error("CopilotService", "Failed to create conversation", error);
      }
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
    signal?: AbortSignal,
  ): Promise<ICopilotResponse> {
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
      const request = this.copilotApi(`/${conversationId}/chat`);
      if (signal) {
        request.option("signal", signal);
      }
      const res = await request.post(requestBody);

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
      if (this.isAbortError(error)) {
        throw error;
      }

      // If the conversation expired or was invalid, reset cache and retry once
      if (this.isConversationNotFound(error)) {
        logger.warn("CopilotService", "Conversation expired, creating new one");
        this.cachedConversation = null;
        const newConversationId = await this.ensureConversation(true, signal);

        const retryRequest = this.copilotApi(`/${newConversationId}/chat`);
        if (signal) {
          retryRequest.option("signal", signal);
        }
        const retryRes = await retryRequest.post(requestBody);

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
    if (this.isAbortError(error)) {
      return "Copilot request was canceled.";
    }
    if (this.isCopilotEndpointUnavailable(error)) {
      return CopilotService.COPILOT_UNAVAILABLE_MESSAGE;
    }
    const statusCode = this.getErrorStatusCode(error);
    if (statusCode === 401 || statusCode === 403) {
      return "You do not have permission or a valid Microsoft 365 Copilot license to use this feature.";
    }
    if (statusCode === 429) {
      return "Copilot is busy. Please try again later.";
    }
    if (statusCode === 400) {
      return "Invalid request to Copilot. Please try rephrasing your input.";
    }
    if (statusCode === 404 || statusCode === 410) {
      return "Copilot conversation expired. Please try again.";
    }
    if (statusCode !== undefined && statusCode >= 500) {
      return "Copilot service is temporarily unavailable. Please try again later.";
    }
    return "An unexpected error occurred while communicating with Copilot.";
  }

  private beginOperation(): AbortController {
    const controller = new AbortController();
    this.activeAbortControllers.add(controller);
    return controller;
  }

  private endOperation(controller: AbortController): void {
    this.activeAbortControllers.delete(controller);
  }

  private createErrorResponse(error: unknown): ICopilotResponse {
    if (this.isAbortError(error)) {
      return {
        content: "",
        isError: true,
        isCancelled: true,
      };
    }

    return {
      content: "",
      isError: true,
      errorMessage: this.getFriendlyErrorMessage(error),
    };
  }

  private parseBooleanLine(line: string): boolean | undefined {
    const separatorIndex = line.indexOf(":");
    if (separatorIndex === -1) {
      return undefined;
    }
    const value = line.substring(separatorIndex + 1).trim().toLowerCase();
    if (value === "yes" || value === "true") {
      return true;
    }
    if (value === "no" || value === "false") {
      return false;
    }
    return undefined;
  }

  private getErrorStatusCode(error: unknown): number | undefined {
    if (error && typeof error === "object") {
      const maybeStatus = (error as { statusCode?: unknown; status?: unknown });
      if (typeof maybeStatus.statusCode === "number") {
        return maybeStatus.statusCode;
      }
      if (typeof maybeStatus.status === "number") {
        return maybeStatus.status;
      }
    }

    const message = error instanceof Error ? error.message : String(error);
    const statusMatch = message.match(/\b([45]\d{2})\b/);
    return statusMatch ? Number(statusMatch[1]) : undefined;
  }

  private getErrorCode(error: unknown): string | undefined {
    if (!error || typeof error !== "object") {
      return undefined;
    }
    const maybeCode = (error as { code?: unknown }).code;
    return typeof maybeCode === "string" ? maybeCode : undefined;
  }

  private isAbortError(error: unknown): boolean {
    const message = error instanceof Error ? error.message : String(error);
    if (typeof DOMException !== "undefined" && error instanceof DOMException) {
      return error.name === "AbortError";
    }
    return (
      message.includes("AbortError") ||
      message.includes("aborted") ||
      message.includes("The user aborted a request")
    );
  }

  private isConversationNotFound(error: unknown): boolean {
    const statusCode = this.getErrorStatusCode(error);
    if (statusCode === 404 || statusCode === 410) {
      return true;
    }

    const code = this.getErrorCode(error)?.toLowerCase();
    if (code === "notfound" || code === "itemnotfound") {
      return true;
    }

    const message = error instanceof Error ? error.message.toLowerCase() : "";
    return message.includes("notfound") || message.includes("not found");
  }

  private isCopilotEndpointUnavailable(error: unknown): boolean {
    const message = error instanceof Error ? error.message.toLowerCase() : "";
    const code = this.getErrorCode(error)?.toLowerCase();
    const statusCode = this.getErrorStatusCode(error);

    return (
      message.includes("resource not found for the segment 'beta'") ||
      message.includes("resource not found for the segment beta") ||
      message.includes("copilot api is not available in this environment") ||
      (statusCode === 404 &&
        (message.includes("/beta/") || message.includes("/copilot/"))) ||
      code === "resourceNotFound".toLowerCase()
    );
  }
}
