import * as React from "react";
import { tokens, Button } from "@fluentui/react-components"; // Keep tokens for gap
import { Link24Regular } from "@fluentui/react-icons";
import { IAlertItem, IQuickAction } from "../Alerts/IAlerts";
import RichMediaAlert from "../Services/RichMediaAlert";
import DescriptionContent from "./DescriptionContent"; // Import DescriptionContent




interface IAlertContentProps {
  item: IAlertItem;
  richMediaEnabled: boolean;
  expanded: boolean;
  stopPropagation: (e: React.MouseEvent) => void;
  handleQuickAction: (action: IQuickAction) => void;
}

const AlertContent: React.FC<IAlertContentProps> = React.memo(({ item, richMediaEnabled, expanded, stopPropagation, handleQuickAction }) => {
  const renderQuickActions = React.useCallback(() => {
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
  }, [item.quickActions, stopPropagation, handleQuickAction, item.Id]);

  return (
    <>
      {item.richMedia && richMediaEnabled && (
        <div onClick={stopPropagation}>
          <RichMediaAlert media={item.richMedia} expanded={expanded} />
        </div>
      )}
      {expanded && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
          {item.richMedia && richMediaEnabled && (
            <div style={{ marginTop: '16px' }}>
              <RichMediaAlert media={item.richMedia} expanded={true} />
            </div>
          )}
          {renderQuickActions()}
        </div>
      )}
    </>
  );
});

export default AlertContent;
