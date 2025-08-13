import * as React from 'react';
import {
  Card,
  CardHeader,
  CardPreview,
  Text,
  Button,
  MessageBar,
  Spinner,
  Badge,
  tokens
} from '@fluentui/react-components';
import { 
  List24Regular, 
  Add24Regular, 
  CheckmarkCircle24Filled,
  ErrorCircle24Filled,
  Warning24Filled,
  Globe24Regular,
  Building24Regular,
  Home24Regular
} from '@fluentui/react-icons';
import { useLocalization } from '../Hooks/useLocalization';
import { SiteContextService, ISiteInfo, IAlertListStatus } from '../Services/SiteContextService';
import styles from './ListManagement.module.scss';

export interface IListManagementProps {
  siteContextService: SiteContextService;
  onListCreated?: (siteId: string) => void;
  className?: string;
}

const ListManagement: React.FC<IListManagementProps> = ({
  siteContextService,
  onListCreated,
  className
}) => {
  const { getString } = useLocalization();
  const [sites, setSites] = React.useState<ISiteInfo[]>([]);
  const [listStatuses, setListStatuses] = React.useState<{ [siteId: string]: IAlertListStatus }>({});
  const [loading, setLoading] = React.useState(true);
  const [creatingList, setCreatingList] = React.useState<string | null>(null);
  const [message, setMessage] = React.useState<{ type: 'success' | 'error'; text: string } | null>(null);

  React.useEffect(() => {
    loadSiteInformation();
  }, [siteContextService]);

  const loadSiteInformation = async () => {
    try {
      setLoading(true);
      
      // Get site hierarchy
      const siteHierarchy = siteContextService.getSitesHierarchy();
      setSites(siteHierarchy);

      // Check list status for each site
      const statuses: { [siteId: string]: IAlertListStatus } = {};
      for (const site of siteHierarchy) {
        try {
          statuses[site.id] = await siteContextService.getAlertListStatus(site.id);
        } catch (error) {
          statuses[site.id] = {
            exists: false,
            canAccess: false,
            canCreate: false,
            error: error.message
          };
        }
      }
      setListStatuses(statuses);
    } catch (error) {
      setMessage({
        type: 'error',
        text: getString('FailedToLoadSiteInformation') || 'Failed to load site information'
      });
    } finally {
      setLoading(false);
    }
  };

  const handleCreateList = async (siteId: string, siteName: string) => {
    try {
      setCreatingList(siteId);
      setMessage(null);

      const success = await siteContextService.createAlertsList(siteId);
      
      if (success) {
        setMessage({
          type: 'success',
          text: getString('AlertsListCreatedSuccessfully') || `Alerts list created successfully on ${siteName}`
        });

        // Refresh site context and list statuses
        await siteContextService.refresh();
        await loadSiteInformation();
        
        if (onListCreated) {
          onListCreated(siteId);
        }
      } else {
        throw new Error('Creation failed');
      }
    } catch (error) {
      setMessage({
        type: 'error',
        text: error.message || getString('FailedToCreateAlertsList') || `Failed to create alerts list on ${siteName}`
      });
    } finally {
      setCreatingList(null);
    }
  };

  const getSiteIcon = (siteType: string) => {
    switch (siteType) {
      case 'home': return <Home24Regular />;
      case 'hub': return <Building24Regular />;
      default: return <Globe24Regular />;
    }
  };

  const getSiteTypeLabel = (siteType: string) => {
    switch (siteType) {
      case 'home': return getString('HomeSite') || 'Home Site';
      case 'hub': return getString('HubSite') || 'Hub Site';
      default: return getString('CurrentSite') || 'Current Site';
    }
  };

  const getSiteDescription = (siteType: string) => {
    switch (siteType) {
      case 'home': return getString('HomeSiteDescription') || 'Alerts shown on all sites in the tenant';
      case 'hub': return getString('HubSiteDescription') || 'Alerts shown on hub and connected sites';
      default: return getString('CurrentSiteDescription') || 'Alerts shown only on this site';
    }
  };

  const getStatusIcon = (status: IAlertListStatus) => {
    if (status.exists && status.canAccess) {
      return <CheckmarkCircle24Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />;
    } else if (status.exists && !status.canAccess) {
      return <Warning24Filled style={{ color: tokens.colorPaletteYellowForeground1 }} />;
    } else if (!status.exists && status.canCreate) {
      return <Add24Regular style={{ color: tokens.colorNeutralForeground3 }} />;
    } else {
      return <ErrorCircle24Filled style={{ color: tokens.colorPaletteRedForeground1 }} />;
    }
  };

  const getStatusText = (status: IAlertListStatus) => {
    if (status.exists && status.canAccess) {
      return getString('ListExistsAndAccessible') || 'List exists and accessible';
    } else if (status.exists && !status.canAccess) {
      return getString('ListExistsNoAccess') || 'List exists but no access';
    } else if (!status.exists && status.canCreate) {
      return getString('ListNotExistsCanCreate') || 'List not found - can create';
    } else {
      return getString('ListNotExistsCannotCreate') || 'List not found - cannot create';
    }
  };

  if (loading) {
    return (
      <div className={`${styles.listManagement} ${className || ''}`}>
        <Card>
          <CardHeader
            image={<List24Regular />}
            header={<Text weight="semibold">{getString('AlertListsManagement') || 'Alert Lists Management'}</Text>}
          />
          <CardPreview>
            <div style={{ padding: '16px', textAlign: 'center' }}>
              <Spinner label={getString('LoadingSiteInformation') || 'Loading site information...'} />
            </div>
          </CardPreview>
        </Card>
      </div>
    );
  }

  return (
    <div className={`${styles.listManagement} ${className || ''}`}>
      {message && (
        <MessageBar intent={message.type} style={{ marginBottom: '16px' }}>
          {message.text}
        </MessageBar>
      )}

      <Card>
        <CardHeader
          image={<List24Regular />}
          header={<Text weight="semibold">{getString('AlertListsManagement') || 'Alert Lists Management'}</Text>}
          description={
            <Text size={200}>
              {getString('ManageAlertListsDescription') || 'Manage alert lists across your site hierarchy'}
            </Text>
          }
        />
      </Card>

      <div className={styles.sitesGrid}>
        {sites.map(site => {
          const status = listStatuses[site.id];
          if (!status) return null;

          return (
            <Card key={site.id} className={styles.siteCard}>
              <CardHeader
                image={getSiteIcon(site.type)}
                header={
                  <div className={styles.siteHeader}>
                    <Text weight="semibold">{site.name}</Text>
                    <Badge appearance="tint" color="informative">
                      {getSiteTypeLabel(site.type)}
                    </Badge>
                  </div>
                }
                description={<Text size={200}>{getSiteDescription(site.type)}</Text>}
              />
              
              <CardPreview>
                <div className={styles.siteStatus}>
                  <div className={styles.statusInfo}>
                    <div className={styles.statusIndicator}>
                      {getStatusIcon(status)}
                      <Text>{getStatusText(status)}</Text>
                    </div>
                    
                    {status.error && (
                      <Text size={200} style={{ color: tokens.colorPaletteRedForeground1 }}>
                        {status.error}
                      </Text>
                    )}
                  </div>

                  {!status.exists && status.canCreate && (
                    <Button
                      appearance="primary"
                      icon={<Add24Regular />}
                      onClick={() => handleCreateList(site.id, site.name)}
                      disabled={creatingList === site.id}
                    >
                      {creatingList === site.id 
                        ? (getString('CreatingList') || 'Creating...')
                        : (getString('CreateAlertsList') || 'Create Alerts List')
                      }
                    </Button>
                  )}

                  {status.exists && status.canAccess && (
                    <div className={styles.listInfo}>
                      <Text size={200} style={{ color: tokens.colorPaletteGreenForeground1 }}>
                        ✓ {getString('ReadyForAlerts') || 'Ready for alerts'}
                      </Text>
                    </div>
                  )}
                </div>
              </CardPreview>
            </Card>
          );
        })}
      </div>

      <Card className={styles.hierarchyInfo}>
        <CardHeader
          header={<Text weight="semibold">{getString('AlertHierarchy') || 'Alert Display Hierarchy'}</Text>}
        />
        <CardPreview>
          <div className={styles.hierarchyDescription}>
            <Text size={200}>
              {getString('AlertHierarchyDescription') || 
                'Alerts are displayed based on site hierarchy: Home Site alerts appear everywhere, Hub Site alerts appear on hub and connected sites, and Site alerts appear only on the specific site.'}
            </Text>
            
            <div className={styles.hierarchyList}>
              <div className={styles.hierarchyItem}>
                <Home24Regular />
                <Text size={200}><strong>{getString('HomeSite') || 'Home Site'}:</strong> {getString('HomeSiteScope') || 'Shown on all sites'}</Text>
              </div>
              <div className={styles.hierarchyItem}>
                <Building24Regular />
                <Text size={200}><strong>{getString('HubSite') || 'Hub Site'}:</strong> {getString('HubSiteScope') || 'Shown on hub and connected sites'}</Text>
              </div>
              <div className={styles.hierarchyItem}>
                <Globe24Regular />
                <Text size={200}><strong>{getString('CurrentSite') || 'Site'}:</strong> {getString('SiteScope') || 'Shown only on this site'}</Text>
              </div>
            </div>
          </div>
        </CardPreview>
      </Card>
    </div>
  );
};

export default ListManagement;