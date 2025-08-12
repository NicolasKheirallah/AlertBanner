import * as React from "react";
import { Button, Text, tokens } from "@fluentui/react-components";
import {
  Dismiss24Regular,
  ChevronLeft24Regular,
  ChevronRight24Regular,
  EyeOff24Regular,
  Link24Regular
} from "@fluentui/react-icons";
import { IAlertItem } from "../Alerts/IAlerts";
import styles from "./AlertItem.module.scss";

interface IAlertActionsProps {
  item: IAlertItem;
  isCarousel: boolean;
  currentIndex: number;
  totalAlerts: number;
  onNext?: () => void;
  onPrevious?: () => void;
  expanded: boolean;
  toggleExpanded: () => void;
  remove: (id: number) => void;
  hideForever: (id: number) => void;
  stopPropagation: (e: React.MouseEvent) => void;
}

const AlertActions: React.FC<IAlertActionsProps> = React.memo(({
  item,
  isCarousel,
  currentIndex,
  totalAlerts,
  onNext,
  onPrevious,
  expanded,
  toggleExpanded,
  remove,
  hideForever,
  stopPropagation
}) => {
  return (
    <div className={styles.actionSection} onClick={stopPropagation}>
      {item.link && (
        <Button
          appearance="subtle"
          icon={<Link24Regular />}
          onClick={(e) => {
            stopPropagation(e);
            if (item.link?.Url) {
              window.open(item.link.Url, "_blank", "noopener,noreferrer");
            }
          }}
          aria-label={item.link.Description || "Open link"}
          size="small"
        />
      )}
      {isCarousel && totalAlerts > 1 && (
        <>
          <Button
            appearance="subtle"
            icon={<ChevronLeft24Regular />}
            onClick={onPrevious}
            aria-label="Previous Alert"
            size="small"
          />
          <div className={styles.carouselCounter}>
            <Text size={200} weight="medium" style={{ color: tokens.colorNeutralForeground2 }}>
              {currentIndex} of {totalAlerts}
            </Text>
          </div>
          <Button
            appearance="subtle"
            icon={<ChevronRight24Regular />}
            onClick={onNext}
            aria-label="Next Alert"
            size="small"
          />
          <div className={styles.divider} />
        </>
      )}
      
      <Button
        appearance="subtle"
        icon={<Dismiss24Regular />}
        onClick={() => remove(item.Id)}
        aria-label="Dismiss Alert"
        size="small"
        title="Dismiss"
      />
      <Button
        appearance="subtle"
        icon={<EyeOff24Regular />}
        onClick={() => hideForever(item.Id)}
        aria-label="Hide Alert Forever"
        size="small"
        title="Hide Forever"
      />
    </div>
  );
});

export default AlertActions;
