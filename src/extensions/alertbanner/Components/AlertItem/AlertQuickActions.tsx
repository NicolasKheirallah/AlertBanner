import * as React from "react";
import { Button, tokens } from "@fluentui/react-components";
import { Link24Regular } from "@fluentui/react-icons";
import { IAlertItem, IQuickAction } from "../Alerts/IAlerts";

interface IAlertQuickActionsProps {
  item: IAlertItem;
  handleQuickAction: (action: IQuickAction) => void;
  stopPropagation: (e: React.MouseEvent) => void;
}

const AlertQuickActions: React.FC<IAlertQuickActionsProps> = React.memo(({ item, handleQuickAction, stopPropagation }) => {
  if (!item.quickActions?.length) return null;

  return (
    <div
      style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalS }}
      onClick={stopPropagation}
    >
      {item.quickActions.map((action: IQuickAction, index: number) => {
        return (
          <Button
            key={`${item.Id}-action-${index}`}
            appearance="outline"
            size="small"
            icon={<Link24Regular />}
            onClick={(e) => {
              stopPropagation(e);
              handleQuickAction(action);
            }}
          >
            {action.label}
          </Button>
        );
      })}
    </div>
  );
});

export default AlertQuickActions;
