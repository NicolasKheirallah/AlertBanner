import * as React from "react";
import styles from "./ColorPicker.module.scss";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

interface IColorPickerProps {
  label: string;
  value: string;
  onChange: (color: string) => void;
  presetColors?: string[];
  description?: string;
  className?: string;
}

const PresetColorButton: React.FC<{
  color: string;
  isSelected: boolean;
  onSelect: () => void;
}> = ({ color, isSelected, onSelect }) => {
  const colorClassName = getColorTokenClass(color);

  return (
    <button
      type="button"
      className={`${styles.presetColor} ${colorClassName} ${isSelected ? styles.selected : ''}`}
      onClick={onSelect}
      aria-label={Text.format(strings.ColorPickerSelectColorAria, color)}
    />
  );
};

const DEFAULT_PRESET_COLORS = [
  "#0078d4", // SharePoint Blue
  "#107c10", // SharePoint Green  
  "#ff8c00", // SharePoint Orange
  "#d13438", // SharePoint Red
  "#5c2d91", // SharePoint Purple
  "#00bcf2", // SharePoint Cyan
  "#ca5010", // SharePoint Dark Orange
  "#8764b8", // SharePoint Light Purple
  "#00b7c3", // SharePoint Teal
  "#bad80a", // SharePoint Lime
  "#ffaa44", // SharePoint Amber
  "#e81123", // SharePoint Error Red
  "#767676", // SharePoint Gray
  "#323130", // SharePoint Dark Gray
  "#000000", // Black
  "#ffffff"  // White
];

const normalizeColorToken = (color: string): string =>
  (color || "").toLowerCase().replace("#", "").replace(/[^a-f0-9]/g, "");
const colorClassMap = styles as unknown as Record<string, string>;
const getColorTokenClass = (color: string): string =>
  colorClassMap[`colorToken${normalizeColorToken(color)}`] ||
  colorClassMap["colorTokenUnknown"] ||
  "";

const ColorPicker: React.FC<IColorPickerProps> = ({
  label,
  value,
  onChange,
  presetColors = DEFAULT_PRESET_COLORS,
  description,
  className
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [customColor, setCustomColor] = React.useState(value);
  const containerRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    setCustomColor(value);
  }, [value]);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    if (isOpen) {
      document.addEventListener('click', handleClickOutside);
    }

    return () => {
      document.removeEventListener('click', handleClickOutside);
    };
  }, [isOpen]);

  const handleColorSelect = React.useCallback((color: string) => {
    onChange(color);
    setCustomColor(color);
  }, [onChange]);

  const isValidColor = React.useCallback((color: string): boolean => {
    if (!color) return false;
    
    const hexPattern = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
    if (hexPattern.test(color)) return true;
    
    const cssColors = [
      'red', 'green', 'blue', 'white', 'black', 'yellow', 'orange', 'purple',
      'pink', 'brown', 'gray', 'grey', 'cyan', 'magenta', 'lime', 'maroon',
      'navy', 'olive', 'teal', 'silver', 'aqua', 'fuchsia'
    ];
    
    return cssColors.includes(color.toLowerCase());
  }, []);

  const selectedColorClassName =
    getColorTokenClass(value);

  return (
    <div className={`${styles.field} ${className || ''}`} ref={containerRef}>
      <label className={styles.label}>
        {label}
      </label>

      {description && (
        <div className={styles.description}>
          {description}
        </div>
      )}

      <div className={styles.colorPickerContainer}>
        <button
          type="button"
          className={styles.colorButton}
          onClick={() => setIsOpen(!isOpen)}
          aria-label={Text.format(strings.ColorPickerSelectedColorAria, value)}
        >
          <div className={styles.colorPreview}>
            <div className={`${styles.colorSwatch} ${selectedColorClassName}`} />
            <span className={styles.colorValue}>{value}</span>
          </div>
        </button>

        {isOpen && (
          <div className={styles.colorDropdown} onClick={(e) => e.stopPropagation()}>
            <div className={styles.presetColors}>
              <h4>{strings.ColorPickerPresetColorsTitle}</h4>
              <div className={styles.colorGrid}>
                {presetColors.map((color) => (
                  <PresetColorButton
                    key={color}
                    color={color}
                    isSelected={color === value}
                    onSelect={() => {
                      handleColorSelect(color);
                      setIsOpen(false);
                    }}
                  />
                ))}
              </div>
            </div>

            <div className={styles.customColorInputs} onClick={(e) => e.stopPropagation()}>
              <input
                type="color"
                value={isValidColor(customColor) ? customColor : '#0078d4'}
                onChange={(e) => {
                  const newColor = e.target.value;
                  onChange(newColor);
                  setCustomColor(newColor);
                }}
                className={styles.nativeColorPicker}
                title="Choose custom color"
                onClick={(e) => e.stopPropagation()}
              />
              <input
                type="text"
                value={customColor}
                onChange={(e) => {
                  const newColor = e.target.value;
                  setCustomColor(newColor);
                  if (isValidColor(newColor)) {
                    onChange(newColor);
                  }
                }}
                onClick={(e) => e.stopPropagation()}
                className={styles.colorTextInput}
                placeholder="#000000"
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default ColorPicker;
