import * as React from "react";
import { IconButton } from "@fluentui/react";
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
        <IconButton
          onRenderIcon={() => <Link24Regular />}
          onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
            stopPropagation(e);
            if (item.linkUrl) {
              window.open(item.linkUrl, WINDOW_OPEN_CONFIG.TARGET, WINDOW_OPEN_CONFIG.FEATURES);
            }
          }}
          aria-label={item.linkDescription || strings.AlertActionsOpenLink}
          title={item.linkDescription || strings.AlertActionsOpenLink}
          className={styles.iconButton}
        />
      )}
      {isCarousel && totalAlerts > 1 && (
        <>
          <IconButton
            onRenderIcon={() => <ChevronLeft24Regular />}
            onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
              stopPropagation(e);
              onPrevious?.();
            }}
            aria-label={strings.AlertActionsPrevious}
            title={strings.AlertActionsPrevious}
            className={styles.iconButton}
          />
          <div className={styles.carouselCounter}>
            <span className={styles.carouselCounterText}>
              {carouselCounterLabel}
            </span>
          </div>
          <IconButton
            onRenderIcon={() => <ChevronRight24Regular />}
            onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
              stopPropagation(e);
              onNext?.();
            }}
            aria-label={strings.AlertActionsNext}
            title={strings.AlertActionsNext}
            className={styles.iconButton}
          />
          <div className={styles.divider} />
        </>
      )}
      <IconButton
        onRenderIcon={() => <Dismiss24Regular />}
        onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
          stopPropagation(e);
          remove(item.id);
        }}
        aria-label={strings.AlertActionsDismiss}
        title={strings.AlertActionsDismiss}
        className={styles.iconButton}
      />
      {userTargetingEnabled && (
        <IconButton
          onRenderIcon={() => <EyeOff24Regular />}
          onClick={(e: React.MouseEvent<HTMLButtonElement>) => {
            stopPropagation(e);
            hideForever(item.id);
          }}
          aria-label={strings.AlertActionsHideForever}
          title={strings.AlertActionsHideForever}
          className={styles.iconButton}
        />
      )}
    </div>
  );
});

AlertActions.displayName = 'AlertActions';

export default AlertActions;
