import * as React from "react";
import styles from "./ColorPicker.module.scss";

interface IColorPickerProps {
  label: string;
  value: string;
  onChange: (color: string) => void;
  presetColors?: string[];
  description?: string;
  className?: string;
}

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

  // Close dropdown when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  const handleColorSelect = (color: string) => {
    onChange(color);
    setCustomColor(color);
    setIsOpen(false);
  };

  const handleCustomColorChange = (color: string) => {
    setCustomColor(color);
    onChange(color);
  };

  const isValidColor = (color: string): boolean => {
    const testElement = document.createElement('div');
    testElement.style.color = color;
    return testElement.style.color !== '';
  };

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
          style={{ backgroundColor: value }}
          aria-label={`Selected color: ${value}`}
        >
          <div className={styles.colorPreview}>
            <div className={styles.colorSwatch} style={{ backgroundColor: value }} />
            <span className={styles.colorValue}>{value}</span>
          </div>
        </button>

        {isOpen && (
          <div className={styles.colorDropdown}>
            <div className={styles.presetColors}>
              <h4>Preset Colors</h4>
              <div className={styles.colorGrid}>
                {presetColors.map((color) => (
                  <button
                    key={color}
                    type="button"
                    className={`${styles.presetColor} ${color === value ? styles.selected : ''}`}
                    style={{ backgroundColor: color }}
                    onClick={() => handleColorSelect(color)}
                    aria-label={`Select color ${color}`}
                    title={color}
                  />
                ))}
              </div>
            </div>

            <div className={styles.customColor}>
              <h4>Custom Color</h4>
              <div className={styles.customColorInputs}>
                <input
                  type="color"
                  value={customColor}
                  onChange={(e) => handleCustomColorChange(e.target.value)}
                  className={styles.nativeColorPicker}
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
                  className={styles.colorTextInput}
                  placeholder="#000000"
                />
                <button
                  type="button"
                  className={styles.applyButton}
                  onClick={() => handleColorSelect(customColor)}
                  disabled={!isValidColor(customColor)}
                >
                  Apply
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default ColorPicker;