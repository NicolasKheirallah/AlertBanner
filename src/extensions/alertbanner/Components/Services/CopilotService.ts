import { MSGraphClientV3 } from "@microsoft/sp-http";
import { logger } from "./LoggerService";

export interface ICopilotResponse {
  content: string;
  isError: boolean;
  isCancelled?: boolean;
  errorMessage?: string;
  citations?: ICopilotCitation[];
}

export interface ICopilotCitation {
  title: string;
  url?: string;
  snippet?: string;
}

export type CopilotTone = "Professional" | "Urgent" | "Casual";

export interface ISentimentResult {
  isProfessional: boolean;
  isToneAppropriate: boolean;
  issues: string[];
  status: "green" | "yellow" | "red";
  rawContent: string;
}

interface IConversationMetadata {
  id: string;
  createdDateTime: string;
  status: string;
  turnCount: number;
}

interface ICopilotLocationHint {
  timeZone: string;
  countryOrRegion?: string;
}

// Service for interacting with the Microsoft 365 Copilot Chat API via Microsoft Graph.
// Uses the /beta/copilot/conversations endpoint which is in preview and subject to change.
export class CopilotService {
  private graphClient: MSGraphClientV3;
  private readonly endpoint = "/copilot/conversations";
  private cachedConversation: IConversationMetadata | null = null;
  private activeAbortControllers = new Set<AbortController>();
  private copilotAvailability: "unknown" | "available" | "unavailable" =
    "unknown";

  private static readonly MAX_CONVERSATION_AGE_MS = 30 * 60 * 1000;
  private static readonly COPILOT_UNAVAILABLE_MESSAGE =
    "Copilot API is not available in this environment.";

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  private copilotApi(pathSuffix: string = "") {
    return this.graphClient
      .api(`${this.endpoint}${pathSuffix}`)
      .version("beta");
  }

  private buildLocationHint(): ICopilotLocationHint {
    let timeZone = "UTC";
    try {
      const resolved = Intl.DateTimeFormat().resolvedOptions().timeZone;
      if (typeof resolved === "string" && resolved.trim().length > 0) {
        timeZone = resolved.trim();
      }
    } catch {
    }

    let countryOrRegion: string | undefined;
    if (typeof navigator !== "undefined") {
      const navLanguage =
        typeof navigator.language === "string" ? navigator.language : "";
      if (navLanguage) {
        const regionMatch = navLanguage.match(/[-_]([A-Za-z]{2})$/);
        if (regionMatch?.[1]) {
          countryOrRegion = regionMatch[1].toUpperCase();
        }
      }
    }

    return {
      timeZone,
      countryOrRegion,
    };
  }

  private sanitizeInput(input: string): string {
    return input
      .replace(/\\/g, "\\\\")
      .replace(/"/g, '\\"')
      .replace(/\n/g, " ")
      .replace(/\r/g, "")
      .trim();
  }

  public async generateDraft(
    keywords: string,
    tone: CopilotTone = "Professional",
  ): Promise<ICopilotResponse> {
    return this.generateDraftWithContext(keywords, tone, false);
  }

  public async generateDraftWithContext(
    promptText: string,
    tone: CopilotTone = "Professional",
    isRefinement: boolean = false,
  ): Promise<ICopilotResponse> {
    const abortController = this.beginOperation();
    try {
      const conversationId = await this.ensureConversation(
        false,
        abortController.signal,
      );
      const sanitizedPrompt = this.sanitizeInput(promptText);

      let prompt: string;
      if (isRefinement) {
        prompt = `You are an assistant for a SharePoint Administrator. 
Refine this ${tone.toLowerCase()} alert banner message based on the following instruction:

${sanitizedPrompt}

Requirements:
- Keep it under 200 words
- Suitable for a corporate intranet
- Be direct and actionable
- Return ONLY the refined message text, no explanations or markdown.`;
      } else {
        prompt = `You are an assistant for a SharePoint Administrator. 
Draft a ${tone.toLowerCase()} alert banner message.
Context and keywords: "${sanitizedPrompt}"

Requirements:
- Keep it under 200 words
- Suitable for a corporate intranet
- Be direct and actionable
- Return ONLY the message text, no explanations or markdown.`;
      }

      return await this.sendMessage(
        conversationId,
        prompt,
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
      SUMMARY: one sentence summary of your analysis
      SUGGESTIONS: If there are issues, provide 1-2 specific suggestions to improve the text. Otherwise "None"`;

      return await this.sendMessage(
        conversationId,
        prompt,
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

      const prompt = `Translate the following text to ${targetLanguage}. 
      Only return the translated text, nothing else.
      
      Text: "${sanitizedText}"`;

      return await this.sendMessage(
        conversationId,
        prompt,
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

  public parseSentimentResult(rawContent: string): ISentimentResult {
    const lines = rawContent.split("\n").map((l) => l.trim());
    const result: ISentimentResult = {
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

  public async checkAccess(): Promise<boolean> {
    const abortController = this.beginOperation();
    try {
      if (this.copilotAvailability === "unavailable") {
        return false;
      }
      // Force a fresh API call to avoid stale true from cached conversation
      await this.ensureConversation(true, abortController.signal);
      return true;
    } catch (error) {
      if (this.isAbortError(error)) {
        return false;
      }
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

  public cancelActiveOperation(): void {
    if (this.activeAbortControllers.size === 0) {
      return;
    }
    this.activeAbortControllers.forEach((controller) => controller.abort());
    this.activeAbortControllers.clear();
  }

  public resetConversation(): void {
    this.cachedConversation = null;
  }

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

  private isConversationExpired(): boolean {
    if (!this.cachedConversation) return true;

    const created = new Date(this.cachedConversation.createdDateTime).getTime();
    const age = Date.now() - created;

    return age > CopilotService.MAX_CONVERSATION_AGE_MS;
  }

  private async sendMessage(
    conversationId: string,
    content: string,
    enableWebGrounding: boolean = false,
    signal?: AbortSignal,
  ): Promise<ICopilotResponse> {
    const requestBody: Record<string, unknown> = {
      message: {
        text: content,
      },
      locationHint: this.buildLocationHint(),
    };

    if (enableWebGrounding) {
      requestBody.enableWebGrounding = true;
    }

    try {
      const request = this.copilotApi(`/${conversationId}/chat`);
      if (signal) {
        request.option("signal", signal);
      }
      const res = await request.post(requestBody);

      if (this.cachedConversation) {
        this.cachedConversation.turnCount++;
      }

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

  private extractResponseText(response: Record<string, unknown>): string {
    // Shape 1: { messages: [{ text: "..." }, { text: "response" }] }
    // The Copilot API returns messages array where last item is the AI response
    const messages = response.messages as Array<Record<string, unknown>> | undefined;
    if (messages && messages.length > 0) {
      const lastMessage = messages[messages.length - 1];
      if (lastMessage?.text) {
        return String(lastMessage.text);
      }
    }

    const message = response.message as Record<string, unknown> | undefined;
    if (message?.text) {
      return String(message.text);
    }

    const value = response.value as Array<Record<string, unknown>> | undefined;
    if (value && value.length > 0 && value[0].content) {
      return String(value[0].content);
    }

    if (response.text) {
      return String(response.text);
    }

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
