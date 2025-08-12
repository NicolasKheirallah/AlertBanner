import * as React from "react";
import { tokens, Button } from "@fluentui/react-components";
import { Link24Regular } from "@fluentui/react-icons";
import { IAlertItem, IQuickAction } from "../Alerts/IAlerts";
import RichMediaAlert from "../Services/RichMediaAlert";
import DescriptionContent from "./DescriptionContent";

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
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: tokens.spacingHorizontalS,
          marginTop: tokens.spacingVerticalM
        }}
        onClick={stopPropagation}
      >
        {item.quickActions.map((action: IQuickAction, index: number) => (
          <Button
            key={`${item.Id}-action-${index}`}
            appearance="primary"
            size="small"
            icon={<Link24Regular />}
            onClick={(e) => {
              stopPropagation(e);
              handleQuickAction(action);
            }}
          >
            {action.label}
          </Button>
        ))}
      </div>
    );
  }, [item.quickActions, stopPropagation, handleQuickAction, item.Id]);

  if (!expanded) return null;

  return (
    <div 
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalM,
        paddingTop: tokens.spacingVerticalM,
        borderTop: `1px solid ${tokens.colorNeutralStroke2}`
      }}
      onClick={stopPropagation}
    >
      {/* Enhanced description content */}
      {item.description && (
        <div>
          <DescriptionContent description={item.description} />
        </div>
      )}
      
      {/* Rich media content */}
      {item.richMedia && richMediaEnabled && (
        <div style={{ marginTop: tokens.spacingVerticalS }}>
          <RichMediaAlert media={item.richMedia} expanded={true} />
        </div>
      )}
      
      {/* Quick actions */}
      {renderQuickActions()}
    </div>
  );
});

export default AlertContent;
