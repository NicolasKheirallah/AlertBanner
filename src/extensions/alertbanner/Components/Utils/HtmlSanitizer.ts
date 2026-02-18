import DOMPurifyFactory from "dompurify";
import type { Config as DOMPurifyConfig } from "dompurify";
import { marked } from "marked";
import { logger } from "../Services/LoggerService";
import { SANITIZATION_CONFIG } from "../Utils/AppConstants";

type DomPurifyInstance = ReturnType<typeof DOMPurifyFactory> | null;

const createDomPurifyInstance = (): DomPurifyInstance => {
  if (typeof window === "undefined") {
    logger.warn(
      "HtmlSanitizer",
      "DOMPurify not initialized because window is undefined",
    );
    return null;
  }

  try {
    return DOMPurifyFactory(window as unknown as Window & typeof globalThis);
  } catch (error) {
    logger.warn("HtmlSanitizer", "Failed to initialize DOMPurify", error);
    return null;
  }
};

const DOMPurify = createDomPurifyInstance();

const getConfiguredTrustedDomains = (): string[] => {
  if (typeof window === "undefined") {
    return [];
  }

  const configured = (window as any).__ALERT_BANNER_TRUSTED_DOMAINS;
  if (!Array.isArray(configured)) {
    return [];
  }

  return configured
    .map((domain) => (domain ?? "").toString().toLowerCase().trim())
    .filter((domain) => domain.length > 0);
};

const isTrustedHost = (hostname: string): boolean => {
  const normalizedHost = hostname.toLowerCase();

  const configuredDomains = getConfiguredTrustedDomains();
  if (
    configuredDomains.some(
      (domain) =>
        normalizedHost === domain || normalizedHost.endsWith(`.${domain}`),
    )
  ) {
    return true;
  }

  // Use centralized configuration from AppConstants
  return SANITIZATION_CONFIG.TRUSTED_DOMAINS.some((pattern) =>
    pattern.test(normalizedHost),
  );
};

export class HtmlSanitizer {
  private static instance: HtmlSanitizer;

  private constructor() {
    this.configureDefaults();
  }

  public static getInstance(): HtmlSanitizer {
    if (!HtmlSanitizer.instance) {
      HtmlSanitizer.instance = new HtmlSanitizer();
    }
    return HtmlSanitizer.instance;
  }

  private configureDefaults(): void {
    if (DOMPurify && typeof DOMPurify.addHook === "function") {
      try {
        DOMPurify.addHook("beforeSanitizeElements", (node: Element) => {
          if (node.tagName === "SCRIPT") {
            node.remove();
          }
        });
      } catch (error) {
        logger.warn(
          "HtmlSanitizer",
          "Failed to configure DOMPurify hooks",
          error,
        );
      }
    }
  }

  public sanitizeHtml(html: string, options?: DOMPurifyConfig): string {
    if (!html) return "";

    if (!DOMPurify || typeof DOMPurify.sanitize !== "function") {
      logger.warn(
        "HtmlSanitizer",
        "DOMPurify not available, using fallback HTML escaping",
      );
      return this.escapeHtml(html);
    }

    const config = {
      ALLOWED_TAGS: [
        "div",
        "span",
        "p",
        "br",
        "strong",
        "b",
        "em",
        "i",
        "u",
        "s",
        "strike",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "ul",
        "ol",
        "li",
        "a",
        "img",
        "blockquote",
        "pre",
        "code",
        "table",
        "thead",
        "tbody",
        "tr",
        "td",
        "th",
      ],
      ALLOWED_ATTR: [
        "href",
        "src",
        "alt",
        "title",
        "class",
        "id",
        "style",
        "target",
        "rel",
        "width",
        "height",
      ],
      ALLOW_DATA_ATTR: false,
      FORBID_TAGS: ["script", "object", "embed", "iframe", "form", "input"],
      FORBID_ATTR: ["onerror", "onload", "onclick", "onmouseover"],
      ...options,
    };

    try {
      const sanitized = DOMPurify.sanitize(html, config) as string | Node;
      if (typeof sanitized === "string") {
        return sanitized;
      }
      // Handle Node/DocumentFragment return types
      if (
        sanitized &&
        typeof sanitized === "object" &&
        "toString" in (sanitized as any)
      ) {
        return (sanitized as any).toString();
      }
      return "";
    } catch (error) {
      logger.error(
        "HtmlSanitizer",
        "DOMPurify sanitization failed, using fallback",
        error,
      );
      return this.escapeHtml(html);
    }
  }

  private escapeHtml(html: string): string {
    if (!html) return "";

    const div = document.createElement("div");
    div.textContent = html;
    return div.innerHTML;
  }

  public markdownToHtml(markdown: string): string {
    if (!markdown) return "";
    marked.setOptions({
      breaks: true,
      gfm: true,
    });

    const html = marked(markdown);

    return this.sanitizeHtml(typeof html === "string" ? html : html.toString());
  }

  public sanitizeAlertContent(content: string): string {
    if (!content) return "";

    const isMarkdown = this.isLikelyMarkdown(content);

    if (isMarkdown) {
      return this.markdownToHtml(content);
    } else {
      return this.sanitizeHtml(content, {
        ALLOWED_TAGS: [
          "div",
          "span",
          "p",
          "br",
          "strong",
          "b",
          "em",
          "i",
          "u",
          "ul",
          "ol",
          "li",
          "a",
          "img",
        ],
        ALLOWED_ATTR: [
          "href",
          "target",
          "rel",
          "class",
          "src",
          "alt",
          "title",
          "width",
          "height",
          "style",
        ],
      });
    }
  }

  private isLikelyMarkdown(content: string): boolean {
    const markdownIndicators = [
      /^\s*#{1,6}\s+/,
      /^\s*\*\s+/,
      /^\s*\d+\.\s+/,
      /\*\*.*\*\*/,
      /\*.*\*/,
      /\[.*\]\(.*\)/,
    ];

    return markdownIndicators.some((pattern) => pattern.test(content));
  }

  public sanitizePreviewContent(content: string): string {
    if (!content) return "";

    // Strip images from preview content - only show text
    return this.sanitizeHtml(content, {
      ALLOWED_TAGS: ["strong", "b", "em", "i", "br", "p"],
      ALLOWED_ATTR: [],
      KEEP_CONTENT: true,
    });
  }

  public sanitizeImageUrl(url: string): boolean {
    if (!url) return false;

    try {
      const urlObj = new URL(url);

      // Allow same origin
      if (urlObj.origin === window.location.origin) {
        return true;
      }

      // Allow https URLs from trusted domains
      if (urlObj.protocol === "https:") {
        return isTrustedHost(urlObj.hostname);
      }

      return false;
    } catch {
      return false;
    }
  }
}

export const htmlSanitizer = HtmlSanitizer.getInstance();
