/**
 * Permission Status Component
 * Displays admin consent requirements and permission status
 */

import * as React from "react";
import { 
  MessageBar, 
  MessageBarBody, 
  MessageBarTitle,
  Button,
  Link
} from "@fluentui/react-components";
import { 
  ShieldError24Regular,
  CheckmarkCircle24Regular 
} from "@fluentui/react-icons";
import { PermissionService, GraphPermission } from "../Services/PermissionService";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import * as strings from 'AlertBannerApplicationCustomizerStrings';

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
      <MessageBar intent="info">
        <MessageBarBody>
          {strings.PermissionStatusChecking}
        </MessageBarBody>
      </MessageBar>
    );
  }

  // All permissions granted
  if (missingPermissions.length === 0) {
    return (
      <MessageBar intent="success">
        <MessageBarBody>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <CheckmarkCircle24Regular />
            <span>{strings.PermissionStatusAllGranted}</span>
          </div>
        </MessageBarBody>
      </MessageBar>
    );
  }

  // Missing critical permissions
  const adminConsentUrl = permissionService.getAdminConsentUrl();
  const isAdmin = context.pageContext.legacyPageContext.isSiteAdmin;

  return (
    <MessageBar intent="warning">
      <MessageBarBody>
        <MessageBarTitle>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <ShieldError24Regular />
            {strings.PermissionStatusMissingTitle}
          </div>
        </MessageBarTitle>
        
        <div style={{ marginTop: '8px' }}>
          {!hasWritePermission && (
            <div style={{ marginBottom: '8px' }}>
              <strong>{strings.PermissionStatusSitesWriteRequired}</strong>
              <p style={{ margin: '4px 0', fontSize: '12px' }}>
                {strings.PermissionStatusSitesWriteDescription}
              </p>
            </div>
          )}
          
          {!hasMailPermission && (
            <div style={{ marginBottom: '8px' }}>
              <strong>{strings.PermissionStatusMailRequired}</strong>
              <p style={{ margin: '4px 0', fontSize: '12px' }}>
                {strings.PermissionStatusMailDescription}
              </p>
            </div>
          )}

          <Button 
            appearance="secondary"
            size="small"
            onClick={() => setState(prev => ({ ...prev, showDetails: !prev.showDetails }))}
          >
            {state.showDetails ? strings.PermissionStatusHideDetails : strings.PermissionStatusShowDetails}
          </Button>

          {state.showDetails && (
            <div style={{ 
              marginTop: '8px', 
              padding: '8px', 
              background: 'rgba(255,255,255,0.5)',
              borderRadius: '4px',
              fontSize: '12px'
            }}>
              <p><strong>{strings.PermissionStatusMissingList}:</strong></p>
              <ul style={{ margin: '4px 0' }}>
                {missingPermissions.map(perm => (
                  <li key={perm}>{perm}</li>
                ))}
              </ul>
              
              {isAdmin && (
                <div style={{ marginTop: '8px' }}>
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
                <div style={{ marginTop: '8px' }}>
                  <p>{strings.PermissionStatusContactAdmin}</p>
                </div>
              )}
            </div>
          )}
        </div>
      </MessageBarBody>
    </MessageBar>
  );
};

export default PermissionStatus;
