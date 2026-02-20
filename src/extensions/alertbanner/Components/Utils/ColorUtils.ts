export const getLuminance = (color: string): number => {
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
      return 0;
    }

    [r, g, b] = resolved;
  }

  const toLinear = (val: number): number => {
    val = val / 255;
    return val <= 0.03928 ? val / 12.92 : Math.pow((val + 0.055) / 1.055, 2.4);
  };

  return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
};

export const getContrastText = (bgColor: string): string => {
  const bgLuminance = getLuminance(bgColor);

  if (bgLuminance > 0.3) {
    return '#323130';
  } else {
    return '#ffffff';
  }
};

export const getContrastRatio = (color1: string, color2: string): number => {
  const lum1 = getLuminance(color1);
  const lum2 = getLuminance(color2);
  const lighter = Math.max(lum1, lum2);
  const darker = Math.min(lum1, lum2);
  return (lighter + 0.05) / (darker + 0.05);
};

export const meetsWCAGAA = (bgColor: string, textColor: string): boolean => {
  return getContrastRatio(bgColor, textColor) >= 4.5;
};

export const meetsWCAGAAA = (bgColor: string, textColor: string): boolean => {
  return getContrastRatio(bgColor, textColor) >= 7.0;
};
