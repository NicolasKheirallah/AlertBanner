import * as React from "react";
import { Button, Text, tokens } from "@fluentui/react-components";
import {
  Dismiss24Regular,
  ChevronLeft24Regular,
  ChevronRight24Regular,
  EyeOff24Regular,
  Link24Regular
} from "@fluentui/react-icons";
import { IAlertItem } from "../Services/SharePointAlertService";
import styles from "./AlertItem.module.scss";
import { useLocalization } from "../Hooks/useLocalization";
import { WINDOW_OPEN_CONFIG } from "../Utils/AppConstants";

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
  const { getString } = useLocalization();
  const carouselCounterLabel = React.useMemo(
    () => getString('AlertCarouselCounter', currentIndex, totalAlerts),
    [currentIndex, totalAlerts, getString]
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
          aria-label={item.linkDescription || getString('AlertActionsOpenLink')}
          size="small"
          title={item.linkDescription || getString('AlertActionsOpenLink')}
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
            aria-label={getString('AlertActionsPrevious')}
            size="small"
            title={getString('AlertActionsPrevious')}
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
            aria-label={getString('AlertActionsNext')}
            size="small"
            title={getString('AlertActionsNext')}
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
        aria-label={getString('AlertActionsDismiss')}
        size="small"
        title={getString('AlertActionsDismiss')}
      />
      {userTargetingEnabled && (
        <Button
          appearance="subtle"
          icon={<EyeOff24Regular />}
          onClick={(e: React.MouseEvent) => {
            stopPropagation(e);
            hideForever(item.id);
          }}
          aria-label={getString('AlertActionsHideForever')}
          size="small"
          title={getString('AlertActionsHideForever')}
        />
      )}
    </div>
  );
});

AlertActions.displayName = 'AlertActions';

export default AlertActions;
