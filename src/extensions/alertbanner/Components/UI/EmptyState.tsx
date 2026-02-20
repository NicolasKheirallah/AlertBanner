/**
 * EmptyState Component
 * 
 * A reusable empty state component following Microsoft 365 first-party design patterns.
 * Uses Fluent UI v8 components for consistent styling and accessibility.
 * 
 * @example
 * <EmptyState
 *   icon="AlertSolid"
 *   title="No alerts found"
 *   description="Get started by creating your first alert"
 *   actionText="Create Alert"
 *   onAction={() => setActiveTab('create')}
 * />
 */

import * as React from "react";
import {
  Stack,
  Text,
  PrimaryButton,
  Icon,
  IStackTokens,
} from "@fluentui/react";
import styles from "./EmptyState.module.scss";

const stackTokens: IStackTokens = {
  childrenGap: 0,
};

const iconStyles = {
  root: {
    fontSize: 48,
    lineHeight: 48,
    color: "#605e5c",
  },
};

/**
 * EmptyState Component
 * 
 * Renders a centered empty state with icon, title, description, and optional action button.
 * Follows Microsoft 365 first-party design guidelines.
 */
const EmptyState: React.FC<{
  icon?: string;
  title: string;
  description: string;
  actionText?: string;
  onAction?: () => void;
  className?: string;
  dataTestId?: string;
}> = ({
  icon,
  title,
  description,
  actionText,
  onAction,
  className,
  dataTestId,
}) => {
  return (
    <Stack
      className={`${styles.emptyState} ${className || ""}`}
      horizontalAlign="center"
      verticalAlign="center"
      tokens={stackTokens}
      data-testid={dataTestId}
    >
      {/* Icon - 48px in neutral gray */}
      {icon && (
        <div className={styles.emptyStateIcon} aria-hidden="true">
          <Icon iconName={icon} styles={{ root: iconStyles }} />
        </div>
      )}

      {/* Title - large variant, semibold */}
      <Text
        variant="large"
        styles={{
          root: {
            fontWeight: 600,
            color: "#323130",
            marginBottom: 8,
          },
        }}
      >
        {title}
      </Text>

      {/* Description - regular, maxWidth 400px, neutral color */}
      <Text
        styles={{
          root: {
            color: "#605e5c",
            maxWidth: 400,
            marginBottom: actionText ? 24 : 0,
            lineHeight: 20,
          },
        }}
      >
        {description}
      </Text>

      {/* Action Button - only render if actionText is provided */}
      {actionText && onAction && (
        <div className={styles.emptyStateAction}>
          <PrimaryButton onClick={onAction}>{actionText}</PrimaryButton>
        </div>
      )}
    </Stack>
  );
};

export default EmptyState;
