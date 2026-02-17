import * as React from "react";
import { 
  MessageBar, 
  MessageBarType,
  DefaultButton,
  Link
} from "@fluentui/react";
import { 
  ShieldError24Regular,
  CheckmarkCircle24Regular 
} from "@fluentui/react-icons";
import { PermissionService, GraphPermission } from "../Services/PermissionService";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import styles from './PermissionStatus.module.scss';

interface IPermissionStatusProps {
  context: ApplicationCustomizerContext;
}

interface IPermissionState {
  isLoading: boolean;
  permissions: Map<GraphPermission, boolean>;
  showDetails: boolean;
}

export const PermissionStatus: React.FC<IPermissionStatusProps> = ({ context }) => {
  const [state, setState] = React.useState<IPermissionState>({
    isLoading: true,
    permissions: new Map(),
    showDetails: false
  });

  const permissionService = React.useMemo(() => 
    PermissionService.getInstance(context), 
    [context]
  );

  React.useEffect(() => {
    const checkPermissions = async () => {
      try {
        const results = await permissionService.validateAllPermissions();
        const permissionMap = new Map<GraphPermission, boolean>();
        results.forEach(r => permissionMap.set(r.scope, r.granted));
        
        setState(prev => ({
          ...prev,
          isLoading: false,
          permissions: permissionMap
        }));
      } catch (error) {
        setState(prev => ({ ...prev, isLoading: false }));
      }
    };

    checkPermissions();
  }, [permissionService]);

  const hasWritePermission = state.permissions.get(GraphPermission.SitesReadWriteAll) || false;
  const hasMailPermission = state.permissions.get(GraphPermission.MailSend) || false;
  const missingPermissions = Array.from(state.permissions.entries())
    .filter(([_, granted]) => !granted)
    .map(([scope]) => scope);

  if (state.isLoading) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        {strings.PermissionStatusChecking}
      </MessageBar>
    );
  }

  // All permissions granted
  if (missingPermissions.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.success}>
        <div className={styles.inlineRow}>
          <CheckmarkCircle24Regular />
          <span>{strings.PermissionStatusAllGranted}</span>
        </div>
      </MessageBar>
    );
  }

  // Missing critical permissions
  const adminConsentUrl = permissionService.getAdminConsentUrl();
  const isAdmin = context.pageContext.legacyPageContext.isSiteAdmin;

  return (
    <MessageBar messageBarType={MessageBarType.warning}>
      <div className={`${styles.inlineRow} ${styles.warningTitle}`}>
        <ShieldError24Regular />
        {strings.PermissionStatusMissingTitle}
      </div>
      
      <div className={styles.spacingTop8}>
        {!hasWritePermission && (
          <div className={styles.spacingBottom8}>
            <strong>{strings.PermissionStatusSitesWriteRequired}</strong>
            <p className={styles.smallParagraph}>
              {strings.PermissionStatusSitesWriteDescription}
            </p>
          </div>
        )}
        
        {!hasMailPermission && (
          <div className={styles.spacingBottom8}>
            <strong>{strings.PermissionStatusMailRequired}</strong>
            <p className={styles.smallParagraph}>
              {strings.PermissionStatusMailDescription}
            </p>
          </div>
        )}

        <DefaultButton 
          onClick={() => setState(prev => ({ ...prev, showDetails: !prev.showDetails }))}
          className={styles.toggleButton}
        >
          {state.showDetails ? strings.PermissionStatusHideDetails : strings.PermissionStatusShowDetails}
        </DefaultButton>

        {state.showDetails && (
          <div className={styles.detailsPanel}>
            <p><strong>{strings.PermissionStatusMissingList}:</strong></p>
            <ul className={styles.listCompact}>
              {missingPermissions.map(perm => (
                <li key={perm}>{perm}</li>
              ))}
            </ul>
            
            {isAdmin && (
              <div className={styles.spacingTop8}>
                <p>{strings.PermissionStatusAdminAction}</p>
                <Link 
                  href={adminConsentUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  {strings.PermissionStatusGrantPermissions}
                </Link>
              </div>
            )}
            
            {!isAdmin && (
              <div className={styles.spacingTop8}>
                <p>{strings.PermissionStatusContactAdmin}</p>
              </div>
            )}
          </div>
        )}
        </div>
    </MessageBar>
  );
};

export default PermissionStatus;
