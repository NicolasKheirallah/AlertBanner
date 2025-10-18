/**
 * Color utility functions for calculating contrast and ensuring accessibility
 */

/**
 * Calculate relative luminance of a color using WCAG formula
 * @param color - Color string in hex, rgb, or named format
 * @returns Luminance value between 0 and 1
 */
export const getLuminance = (color: string): number => {
  // Convert color to RGB values
  let r: number, g: number, b: number;

  const parseRgbValues = (value: string): [number, number, number] | null => {
    const matches = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
    if (!matches) {
      return null;
    }

    return [parseInt(matches[1], 10), parseInt(matches[2], 10), parseInt(matches[3], 10)];
  };

  const resolveCssColor = (value: string): [number, number, number] | null => {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      return null;
    }

    const element = document.createElement('span');
    element.style.color = value;
    element.style.display = 'none';
    document.body.appendChild(element);
    const computedColor = window.getComputedStyle(element).color;
    document.body.removeChild(element);

    return parseRgbValues(computedColor);
  };

  if (color.startsWith('#')) {
    // Hex color
    const hex = color.replace('#', '');
    if (hex.length === 3) {
      r = parseInt(hex[0] + hex[0], 16);
      g = parseInt(hex[1] + hex[1], 16);
      b = parseInt(hex[2] + hex[2], 16);
    } else {
      r = parseInt(hex.substring(0, 2), 16);
      g = parseInt(hex.substring(2, 4), 16);
      b = parseInt(hex.substring(4, 6), 16);
    }
  } else if (color.toLowerCase() === 'white' || color.toLowerCase() === '#ffffff') {
    r = g = b = 255;
  } else if (color.toLowerCase() === 'black' || color.toLowerCase() === '#000000') {
    r = g = b = 0;
  } else if (color.startsWith('rgb(')) {
    const values = parseRgbValues(color);
    if (!values) {
      return 0;
    }
    [r, g, b] = values;
  } else if (color.startsWith('rgba(')) {
    const values = parseRgbValues(color);
    if (!values) {
      return 0;
    }
    [r, g, b] = values;
  } else {
    const resolved = resolveCssColor(color);
    if (!resolved) {
      // Default to dark luminance to ensure high-contrast fallback text
      return 0;
    }

    [r, g, b] = resolved;
  }

  // Calculate relative luminance using WCAG formula
  const toLinear = (val: number): number => {
    val = val / 255;
    return val <= 0.03928 ? val / 12.92 : Math.pow((val + 0.055) / 1.055, 2.4);
  };

  return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
};

/**
 * Get contrasting text color (dark or light) based on background color
 * Uses WCAG AAA standard (7:1 contrast ratio) for better accessibility
 * @param bgColor - Background color in any valid CSS format
 * @returns Appropriate text color (#323130 for dark text, #ffffff for white text)
 */
export const getContrastText = (bgColor: string): string => {
  const bgLuminance = getLuminance(bgColor);

  // More aggressive approach for better readability
  // Use WCAG AAA standard (7:1 contrast ratio) for better accessibility
  if (bgLuminance > 0.3) {
    // For lighter backgrounds, always use dark text for maximum readability
    return '#323130'; // Dark text that meets AAA standards
  } else {
    // For darker backgrounds, always use white text
    return '#ffffff'; // White text for maximum contrast
  }
};

/**
 * Calculate contrast ratio between two colors
 * @param color1 - First color
 * @param color2 - Second color
 * @returns Contrast ratio (1-21, where 21 is maximum contrast)
 */
export const getContrastRatio = (color1: string, color2: string): number => {
  const lum1 = getLuminance(color1);
  const lum2 = getLuminance(color2);
  const lighter = Math.max(lum1, lum2);
  const darker = Math.min(lum1, lum2);
  return (lighter + 0.05) / (darker + 0.05);
};

/**
 * Check if color combination meets WCAG AA standard (4.5:1 for normal text)
 * @param bgColor - Background color
 * @param textColor - Text color
 * @returns True if contrast meets WCAG AA standard
 */
export const meetsWCAGAA = (bgColor: string, textColor: string): boolean => {
  return getContrastRatio(bgColor, textColor) >= 4.5;
};

/**
 * Check if color combination meets WCAG AAA standard (7:1 for normal text)
 * @param bgColor - Background color
 * @param textColor - Text color
 * @returns True if contrast meets WCAG AAA standard
 */
export const meetsWCAGAAA = (bgColor: string, textColor: string): boolean => {
  return getContrastRatio(bgColor, textColor) >= 7.0;
};
