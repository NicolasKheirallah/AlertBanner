
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
      {icon && (
        <div className={styles.emptyStateIcon} aria-hidden="true">
          <Icon iconName={icon} styles={{ root: iconStyles }} />
        </div>
      )}

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

      {actionText && onAction && (
        <div className={styles.emptyStateAction}>
          <PrimaryButton onClick={onAction}>{actionText}</PrimaryButton>
        </div>
      )}
    </Stack>
  );
};

export default EmptyState;
