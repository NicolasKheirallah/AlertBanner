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
import { WINDOW_OPEN_CONFIG } from "../Utils/AppConstants";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';

interface IAlertActionsProps {
  item: IAlertItem;
  isCarousel: boolean;
  currentIndex: number;
  totalAlerts: number;
  onNext?: () => void;
  onPrevious?: () => void;
  expanded: boolean;
  toggleExpanded: () => void;
  remove: (id: string) => void;
  hideForever: (id: string) => void;
  stopPropagation: (e: React.MouseEvent) => void;
  userTargetingEnabled: boolean;
}

const AlertActions: React.FC<IAlertActionsProps> = React.memo(({
  item,
  isCarousel,
  currentIndex,
  totalAlerts,
  onNext,
  onPrevious,
  remove,
  hideForever,
  stopPropagation,
  userTargetingEnabled
}) => {
  const carouselCounterLabel = React.useMemo(
    () => CoreText.format(strings.AlertCarouselCounter, currentIndex.toString(), totalAlerts.toString()),
    [currentIndex, totalAlerts]
  );

  return (
    <div className={styles.actionSection}>
      {item.linkUrl && (
        <Button
          appearance="subtle"
          icon={<Link24Regular />}
          onClick={(e: React.MouseEvent) => {
            stopPropagation(e);
            if (item.linkUrl) {
              window.open(item.linkUrl, WINDOW_OPEN_CONFIG.TARGET, WINDOW_OPEN_CONFIG.FEATURES);
            }
          }}
          aria-label={item.linkDescription || strings.AlertActionsOpenLink}
          size="small"
          title={item.linkDescription || strings.AlertActionsOpenLink}
        />
      )}
      {isCarousel && totalAlerts > 1 && (
        <>
          <Button
            appearance="subtle"
            icon={<ChevronLeft24Regular />}
            onClick={(e: React.MouseEvent) => {
              stopPropagation(e);
              onPrevious?.();
            }}
            aria-label={strings.AlertActionsPrevious}
            size="small"
            title={strings.AlertActionsPrevious}
          />
          <div className={styles.carouselCounter}>
            <Text size={200} weight="medium" style={{ color: tokens.colorNeutralForeground2 }}>
              {carouselCounterLabel}
            </Text>
          </div>
          <Button
            appearance="subtle"
            icon={<ChevronRight24Regular />}
            onClick={(e: React.MouseEvent) => {
              stopPropagation(e);
              onNext?.();
            }}
            aria-label={strings.AlertActionsNext}
            size="small"
            title={strings.AlertActionsNext}
          />
          <div className={styles.divider} />
        </>
      )}
      <Button
        appearance="subtle"
        icon={<Dismiss24Regular />}
        onClick={(e: React.MouseEvent) => {
          stopPropagation(e);
          remove(item.id);
        }}
        aria-label={strings.AlertActionsDismiss}
        size="small"
        title={strings.AlertActionsDismiss}
      />
      {userTargetingEnabled && (
        <Button
          appearance="subtle"
          icon={<EyeOff24Regular />}
          onClick={(e: React.MouseEvent) => {
            stopPropagation(e);
            hideForever(item.id);
          }}
          aria-label={strings.AlertActionsHideForever}
          size="small"
          title={strings.AlertActionsHideForever}
        />
      )}
    </div>
  );
});

AlertActions.displayName = 'AlertActions';

export default AlertActions;
