import * as React from "react";
import { AlertPriority } from "../Alerts/IAlerts";
import {
  Info24Regular,
  Warning24Regular,
  ErrorCircle24Regular,
  CheckmarkCircle24Regular
} from "@fluentui/react-icons";

export const PRIORITY_COLORS = {
  critical: { start: '#C0392B', end: '#E74C3C' }, // Dark Red to Brighter Red
  high: { start: '#D35400', end: '#F39C12' },     // Dark Orange to Golden Orange
  medium: { start: '#2980B9', end: '#3498DB' },    // Dark Blue to Sky Blue
  low: { start: '#27AE60', end: '#2ECC71' }      // Forest Green to Emerald Green
} as const;

export const getPriorityIcon = (priority: AlertPriority): React.ReactElement => {
  const iconColor = PRIORITY_COLORS[priority as keyof typeof PRIORITY_COLORS]?.start || PRIORITY_COLORS.low.start; // Use the start color for the icon
  const iconStyle = { width: 20, height: 20, color: iconColor };
  
  switch (priority) {
    case "critical":
      return <ErrorCircle24Regular style={iconStyle} />;
    case "high":
      return <Warning24Regular style={iconStyle} />;
    case "medium":
      return <Info24Regular style={iconStyle} />;
    case "low":
    default:
      return <CheckmarkCircle24Regular style={iconStyle} />;
  }
};