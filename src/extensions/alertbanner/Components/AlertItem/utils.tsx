import * as React from "react";
import { AlertPriority } from "../Alerts/IAlerts";
import * as FluentIcons from "@fluentui/react-icons";
import styles from "./AlertItem.module.scss";

type IconComponent = React.ComponentType<any>;

const normalizeIconToken = (value: string): string =>
  (value || "").toLowerCase().replace(/[^a-z0-9]/g, "");

const isIconComponent = (value: unknown): value is IconComponent => {
  if (typeof value === "function") {
    return true;
  }

  // React.forwardRef components are objects with React type metadata.
  if (typeof value === "object" && value !== null) {
    const candidate = value as { $$typeof?: unknown; render?: unknown };
    return !!candidate.$$typeof || typeof candidate.render === "function";
  }

  return false;
};

const FLUENT_REGULAR_ICON_MAP: Map<string, IconComponent> = (() => {
  const map = new Map<string, IconComponent>();

  Object.entries(FluentIcons).forEach(([name, icon]) => {
    if (!name.endsWith("24Regular")) {
      return;
    }

    if (!isIconComponent(icon)) {
      return;
    }

    const baseName = name.slice(0, -("24Regular".length));
    map.set(normalizeIconToken(baseName), icon as IconComponent);
  });

  return map;
})();

const ICON_ALIAS_MAP = new Map<string, string>([
  ["error", normalizeIconToken("ErrorCircle")],
  ["checkmark", normalizeIconToken("CheckmarkCircle")],
  ["announcement", normalizeIconToken("Megaphone")],
  ["maintenance", normalizeIconToken("Settings")],
  ["constructioncone", normalizeIconToken("Settings")],
]);

const resolveFluentIconComponent = (
  iconName: string | undefined,
): IconComponent | undefined => {
  if (!iconName) {
    return undefined;
  }

  const normalized = normalizeIconToken(iconName);
  const aliasResolved = ICON_ALIAS_MAP.get(normalized) || normalized;
  const directMatch = FLUENT_REGULAR_ICON_MAP.get(aliasResolved);
  if (directMatch) {
    return directMatch;
  }

  if (aliasResolved.endsWith("24regular")) {
    const withoutSize = aliasResolved.replace(/24regular$/, "");
    return FLUENT_REGULAR_ICON_MAP.get(withoutSize);
  }

  return undefined;
};

export const ALL_FLUENT_ICON_NAMES: string[] = Array.from(
  new Set(
    Object.keys(FluentIcons)
      .filter((name) => name.endsWith("24Regular"))
      .map((name) => name.replace(/24Regular$/, "")),
  ),
).sort((a, b) => a.localeCompare(b));

const getPriorityIconClassName = (priority: AlertPriority): string => {
  switch (priority) {
    case "critical":
      return styles.priorityIconCritical;
    case "high":
      return styles.priorityIconHigh;
    case "medium":
      return styles.priorityIconMedium;
    case "low":
      return styles.priorityIconLow;
    default:
      return styles.priorityIconLow;
  }
};

const getCommonIconProps = (priority: AlertPriority) => {
  return {
    className: `${styles.priorityIcon} ${getPriorityIconClassName(priority)}`,
    width: "20px",
    height: "20px",
    fill: "currentColor",
  };
};

export const getPriorityIcon = (priority: AlertPriority): React.ReactElement => {
  const commonProps = getCommonIconProps(priority);

  switch (priority) {
    case "critical":
      return <FluentIcons.ErrorCircle24Regular {...commonProps} />;
    case "high":
      return <FluentIcons.Warning24Regular {...commonProps} />;
    case "medium":
      return <FluentIcons.Info24Regular {...commonProps} />;
    case "low":
    default:
      return <FluentIcons.CheckmarkCircle24Regular {...commonProps} />;
  }
};

export const getAlertTypeIcon = (
  iconName: string | undefined,
  priority: AlertPriority,
): React.ReactElement => {
  const commonProps = getCommonIconProps(priority);
  const iconComponent = resolveFluentIconComponent(iconName);

  if (iconComponent) {
    return React.createElement(iconComponent, commonProps);
  }

  return getPriorityIcon(priority);
};

export const ALERT_TYPE_ICON_NAMES: string[] = [
  "Info",
  "Warning",
  "Error",
  "CheckmarkCircle",
  "Megaphone",
  "Settings",
  "Shield",
  "Clock",
  "Alert",
  "Sparkle",
  "Book",
  "Trophy",
];

export const getAlertTypeIconLabel = (iconName: string): string => {
  return iconName
    .replace(/([a-z])([A-Z0-9])/g, "$1 $2")
    .replace(/([0-9])([A-Za-z])/g, "$1 $2")
    .trim();
};
