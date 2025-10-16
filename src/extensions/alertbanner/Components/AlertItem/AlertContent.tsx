import * as React from "react";
import { IAlertItem } from "../Services/SharePointAlertService";
import DescriptionContent from "./DescriptionContent";
import styles from "./AlertItem.module.scss";

interface IAlertContentProps {
  item: IAlertItem;
  expanded: boolean;
  stopPropagation: (e: React.MouseEvent) => void;
}

const AlertContent: React.FC<IAlertContentProps> = React.memo(({ item, expanded, stopPropagation }) => {
  if (!expanded) return null;

  return (
    <div
      className={styles.alertContentContainer}
      onClick={stopPropagation}
    >
      {item.description && <DescriptionContent description={item.description} />}
    </div>
  );
});

AlertContent.displayName = 'AlertContent';

export default AlertContent;
