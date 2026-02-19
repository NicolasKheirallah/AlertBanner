import { logger } from "./LoggerService";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import { StringUtils } from "../Utils/StringUtils";
import { JsonUtils } from "../Utils/JsonUtils";
import {
  VALIDATION_LIMITS,
  REGEX_PATTERNS,
  SANITIZATION_CONFIG,
} from "../Utils/AppConstants";

export interface IValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  sanitizedValue?: unknown;
}

export interface IValidationRule {
  name: string;
  message: string;
  validator: (value: unknown) => boolean;
  sanitizer?: (value: unknown) => unknown;
}

export class ValidationService {
  private static _instance: ValidationService;

  private readonly patterns = {
    email: REGEX_PATTERNS.EMAIL,
    url: REGEX_PATTERNS.URL,
    guid: REGEX_PATTERNS.GUID,
    htmlTag: /<[^>]*>/g,
    script: /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
    maliciousPatterns: [
      /javascript:/i,
      /vbscript:/i,
      /on\w+\s*=/i,
      /data:text\/html/i,
      /eval\s*\(/i,
      /expression\s*\(/i,
    ],
  };

  private constructor() {}

  public static getInstance(): ValidationService {
    if (!ValidationService._instance) {
      ValidationService._instance = new ValidationService();
    }
    return ValidationService._instance;
  }

  public validateAlertTitle(title: string): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (StringUtils.isEmpty(title)) {
      errors.push("Title is required");
      return { isValid: false, errors, warnings };
    }

    const trimmedTitle = StringUtils.trimOrDefault(title);

    if (trimmedTitle.length > VALIDATION_LIMITS.TITLE_MAX_LENGTH) {
      errors.push(
        `Title cannot exceed ${VALIDATION_LIMITS.TITLE_MAX_LENGTH} characters`,
      );
    }

    if (trimmedTitle.length < VALIDATION_LIMITS.TITLE_MIN_LENGTH) {
      warnings.push(
        `Title should be at least ${VALIDATION_LIMITS.TITLE_MIN_LENGTH} characters long`,
      );
    }

    if (this.containsMaliciousContent(trimmedTitle)) {
      errors.push("Title contains potentially malicious content");
    }

    const sanitizedValue = this.sanitizeText(trimmedTitle);

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue,
    };
  }

  public validateAlertDescription(description: string): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!description || typeof description !== "string") {
      errors.push("Description is required and must be a string");
      return { isValid: false, errors, warnings };
    }

    const trimmedDescription = description.trim();

    if (trimmedDescription.length === 0) {
      errors.push("Description cannot be empty");
    }

    if (trimmedDescription.length > VALIDATION_LIMITS.DESCRIPTION_MAX_LENGTH) {
      errors.push(
        `Description cannot exceed ${VALIDATION_LIMITS.DESCRIPTION_MAX_LENGTH} characters`,
      );
    }

    if (trimmedDescription.length < 10) {
      warnings.push(
        "Description should be at least 10 characters long for clarity",
      );
    }

    if (this.containsMaliciousContent(trimmedDescription)) {
      errors.push("Description contains potentially malicious content");
    }

    const sanitizedValue = this.sanitizeHtml(trimmedDescription);

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue,
    };
  }

  public validateUrl(
    url: string,
    requireSecure: boolean = true,
  ): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!url || typeof url !== "string") {
      return { isValid: true, errors, warnings, sanitizedValue: "" };
    }

    const trimmedUrl = url.trim();

    if (trimmedUrl.length === 0) {
      return { isValid: true, errors, warnings, sanitizedValue: "" };
    }

    if (trimmedUrl.length > VALIDATION_LIMITS.URL_MAX_LENGTH) {
      errors.push(
        `URL cannot exceed ${VALIDATION_LIMITS.URL_MAX_LENGTH} characters`,
      );
    }

    if (!this.patterns.url.test(trimmedUrl)) {
      errors.push("URL format is invalid");
    }

    if (requireSecure && !trimmedUrl.startsWith("https://")) {
      errors.push("URL must use HTTPS for security");
    }

    if (this.containsMaliciousContent(trimmedUrl)) {
      errors.push("URL contains potentially malicious content");
    }

    try {
      const urlObj = new URL(trimmedUrl);

      if (urlObj.protocol !== "https:" && urlObj.protocol !== "http:") {
        errors.push("URL must use HTTP or HTTPS protocol");
      }

      const suspiciousDomains = ["bit.ly", "tinyurl.com", "short.link"];
      if (
        suspiciousDomains.some((domain) => urlObj.hostname.includes(domain))
      ) {
        warnings.push(
          "URL uses a URL shortener which may obscure the final destination",
        );
      }
    } catch (urlError) {
      errors.push("URL format is invalid");
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue: trimmedUrl,
    };
  }

  public validateDateRange(
    startDate?: Date | string,
    endDate?: Date | string,
  ): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    let parsedStartDate: Date | null = null;
    let parsedEndDate: Date | null = null;

    if (startDate) {
      parsedStartDate =
        typeof startDate === "string" ? new Date(startDate) : startDate;
      if (isNaN(parsedStartDate.getTime())) {
        errors.push("Start date is invalid");
      }
    }

    if (endDate) {
      parsedEndDate = typeof endDate === "string" ? new Date(endDate) : endDate;
      if (isNaN(parsedEndDate.getTime())) {
        errors.push("End date is invalid");
      }
    }

    if (parsedStartDate && parsedEndDate) {
      if (parsedStartDate >= parsedEndDate) {
        errors.push("End date must be after start date");
      }

      const daysDiff =
        (parsedEndDate.getTime() - parsedStartDate.getTime()) /
        (1000 * 60 * 60 * 24);
      if (daysDiff > 365) {
        warnings.push("Alert duration is longer than one year");
      }
    }

    const now = new Date();
    if (parsedStartDate && parsedStartDate < now) {
      warnings.push("Start date is in the past");
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue: {
        startDate: parsedStartDate,
        endDate: parsedEndDate,
      },
    };
  }

  // Validate JSON data with security checks using JsonUtils - prevents prototype pollution
  public validateJson(
    jsonString: string,
    maxDepth: number = VALIDATION_LIMITS.JSON_MAX_DEPTH,
  ): IValidationResult {
    const result = JsonUtils.parseWithValidation(jsonString, {
      maxDepth,
      checkDangerousKeys: true,
    });

    return {
      isValid: result.success,
      errors: result.errors,
      warnings: [],
      sanitizedValue: result.data,
    };
  }

  public validateEmail(email: string): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!email || typeof email !== "string") {
      return { isValid: true, errors, warnings, sanitizedValue: "" };
    }

    const trimmedEmail = email.trim().toLowerCase();

    if (trimmedEmail.length === 0) {
      return { isValid: true, errors, warnings, sanitizedValue: "" };
    }

    if (!this.patterns.email.test(trimmedEmail)) {
      errors.push("Email format is invalid");
    }

    if (trimmedEmail.length > 320) {
      errors.push("Email address is too long (maximum 320 characters)");
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      sanitizedValue: trimmedEmail,
    };
  }

  private containsMaliciousContent(input: string): boolean {
    return this.patterns.maliciousPatterns.some((pattern) =>
      pattern.test(input),
    );
  }

  private sanitizeText(input: string): string {
    return input
      .trim()
      .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
      .replace(/\s+/g, " ");
  }

  // Sanitize HTML using DOMPurify-based HtmlSanitizer for XSS protection
  private sanitizeHtml(input: string): string {
    const sanitized = htmlSanitizer.sanitizeHtml(input);

    if (sanitized !== input) {
      logger.warn(
        "ValidationService",
        "Potential XSS attempt detected and sanitized",
        {
          original: StringUtils.truncate(input, 100),
          sanitized: StringUtils.truncate(sanitized, 100),
        },
      );
    }

    return sanitized;
  }
}

export const validationService = ValidationService.getInstance();
