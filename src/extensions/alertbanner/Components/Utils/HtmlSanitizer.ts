let DOMPurify: any = null;
try {
  DOMPurify = require('dompurify');
} catch (error) {
  logger.warn('HtmlSanitizer', 'DOMPurify not available in this environment', error);
}

import { marked } from 'marked';
import { logger } from '../Services/LoggerService';

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
    if (DOMPurify && typeof DOMPurify.addHook === 'function') {
      try {
        DOMPurify.addHook('beforeSanitizeElements', (node: Element) => {
          if (node.tagName === 'SCRIPT') {
            node.remove();
          }
        });
      } catch (error) {
        logger.warn('HtmlSanitizer', 'Failed to configure DOMPurify hooks', error);
      }
    }
  }

  public sanitizeHtml(html: string, options?: any): string {
    if (!html) return '';
    
    if (!DOMPurify || typeof DOMPurify.sanitize !== 'function') {
      logger.warn('HtmlSanitizer', 'DOMPurify not available, using fallback HTML escaping');
      return this.escapeHtml(html);
    }

    const config = {
      ALLOWED_TAGS: [
        'div', 'span', 'p', 'br', 'strong', 'b', 'em', 'i', 'u', 's', 'strike',
        'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
        'ul', 'ol', 'li',
        'a', 'img',
        'blockquote', 'pre', 'code',
        'table', 'thead', 'tbody', 'tr', 'td', 'th'
      ],
      ALLOWED_ATTR: [
        'href', 'src', 'alt', 'title', 'class', 'id', 'style',
        'target', 'rel', 'width', 'height'
      ],
      ALLOW_DATA_ATTR: false,
      FORBID_TAGS: ['script', 'object', 'embed', 'iframe', 'form', 'input'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'onmouseover'],
      ...options
    };

    try {
      return DOMPurify.sanitize(html, config);
    } catch (error) {
      logger.error('HtmlSanitizer', 'DOMPurify sanitization failed, using fallback', error);
      return this.escapeHtml(html);
    }
  }

  private escapeHtml(html: string): string {
    if (!html) return '';
    
    const div = document.createElement('div');
    div.textContent = html;
    return div.innerHTML;
  }

  public markdownToHtml(markdown: string): string {
    if (!markdown) return '';
    marked.setOptions({
      breaks: true,
      gfm: true
    });

    const html = marked(markdown);

    return this.sanitizeHtml(html as string);
  }

  public sanitizeAlertContent(content: string): string {
    if (!content) return '';

    const isMarkdown = this.isLikelyMarkdown(content);

    if (isMarkdown) {
      return this.markdownToHtml(content);
    } else {
      return this.sanitizeHtml(content, {
        ALLOWED_TAGS: [
          'div', 'span', 'p', 'br', 'strong', 'b', 'em', 'i', 'u',
          'ul', 'ol', 'li', 'a'
        ],
        ALLOWED_ATTR: ['href', 'target', 'rel', 'class']
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

    return markdownIndicators.some(pattern => pattern.test(content));
  }

  public sanitizePreviewContent(content: string): string {
    if (!content) return '';

    return this.sanitizeHtml(content, {
      ALLOWED_TAGS: ['strong', 'b', 'em', 'i', 'br', 'p'],
      ALLOWED_ATTR: [],
      KEEP_CONTENT: true
    });
  }
}

export const htmlSanitizer = HtmlSanitizer.getInstance();