import * as React from "react";
import { AlertPriority } from "../Alerts/IAlerts";
import { tokens } from "@fluentui/react-components";
import {
  Info24Regular,
  Warning24Regular,
  ErrorCircle24Regular,
  CheckmarkCircle24Regular
} from "@fluentui/react-icons";

export const getPriorityIcon = (priority: AlertPriority): React.ReactElement => {
  const getIconColor = (priority: AlertPriority): string => {
    switch (priority) {
      case "critical": return tokens.colorPaletteRedForeground1;
      case "high": return tokens.colorPaletteDarkOrangeForeground1;
      case "medium": return tokens.colorPaletteBlueForeground2;
      case "low": return tokens.colorPaletteGreenForeground1;
      default: return tokens.colorPaletteGreenForeground1;
    }
  };

  const iconColor = getIconColor(priority);

  const commonProps = {
    style: { color: iconColor },
    width: "20px",
    height: "20px",
    fill: "currentColor"
  };

  switch (priority) {
    case "critical":
      return <ErrorCircle24Regular {...commonProps} />;
    case "high":
      return <Warning24Regular {...commonProps} />;
    case "medium":
      return <Info24Regular {...commonProps} />;
    case "low":
    default:
      return <CheckmarkCircle24Regular {...commonProps} />;
  }
};