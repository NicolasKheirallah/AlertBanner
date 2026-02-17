import * as React from 'react';
import { logger } from '../Services/LoggerService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import {
  Checkbox as FluentCheckbox,
  DefaultButton,
  PrimaryButton,
  MessageBar as FluentMessageBar,
  MessageBarType,
  Spinner,
  Dialog as FluentDialog,
} from "@fluentui/react";
import { 
  List24Regular, 
  Add24Regular, 
  CheckmarkCircle24Filled,
  ErrorCircle24Filled,
  Warning24Filled,
  Globe24Regular,
  Building24Regular,
  Home24Regular,
  LocalLanguage24Regular
} from '@fluentui/react-icons';
import { SiteContextService, ISiteInfo, IAlertListStatus } from '../Services/SiteContextService';
import styles from './ListManagement.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';
import { SUPPORTED_LANGUAGES } from '../Utils/AppConstants';

const formatWithFallback = (value: string | undefined, fallback: string, ...args: Array<string | number>): string => {
  const template = value || fallback;
  if (args.length === 0) {
    return template;
  }
  const formattedArgs = args.map(arg => arg.toString());
  return CoreText.format(template, ...formattedArgs);
};

export interface IListManagementProps {
  siteContextService: SiteContextService;
  onListCreated?: (siteId: string) => void;
  className?: string;
}

// Available languages for selection
const AVAILABLE_LANGUAGES = SUPPORTED_LANGUAGES;

const cx = (...classes: Array<string | undefined | false>): string =>
  classes.filter(Boolean).join(" ");

const Card: React.FC<{ children?: React.ReactNode; className?: string }> = ({ children, className }) => (
  <div className={cx(styles.f2Card, className)}>
    {children}
  </div>
);

const CardHeader: React.FC<{
  image?: React.ReactNode;
  header?: React.ReactNode;
  description?: React.ReactNode;
}> = ({ image, header, description }) => (
  <div className={styles.f2CardHeader}>
    {image && <div>{image}</div>}
    <div className={styles.f2CardHeaderText}>
      {header}
      {description}
    </div>
  </div>
);

const CardPreview: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div>{children}</div>;

const Text: React.FC<{
  children?: React.ReactNode;
  size?: number;
  weight?: "regular" | "medium" | "semibold" | "bold";
  className?: string;
}> = ({ children, size, weight, className }) => (
  <span
    className={cx(
      styles.f2Text,
      size === 100 && styles.f2Text100,
      size === 200 && styles.f2Text200,
      size === 300 && styles.f2Text300,
      size === 400 && styles.f2Text400,
      size === 500 && styles.f2Text500,
      weight === "regular" && styles.f2TextRegular,
      weight === "medium" && styles.f2TextMedium,
      weight === "semibold" && styles.f2TextSemibold,
      weight === "bold" && styles.f2TextBold,
      className,
    )}
  >
    {children}
  </span>
);

const Button: React.FC<{
  children?: React.ReactNode;
  appearance?: "primary" | "secondary" | "subtle";
  icon?: React.ReactNode;
  onClick?: () => void | Promise<void>;
  disabled?: boolean;
  className?: string;
  size?: "small" | "medium" | "large";
}> = ({ children, appearance = "secondary", icon, onClick, disabled, className, size }) => {
  const buttonClassName = cx(
    styles.f2Button,
    appearance === "primary" && styles.f2ButtonPrimary,
    appearance === "subtle" && styles.f2ButtonSubtle,
    size === "small" && styles.f2ButtonSmall,
    className,
  );

  const commonProps = {
    onRenderIcon: icon ? () => <>{icon}</> : undefined,
    onClick: onClick as any,
    disabled,
    className: buttonClassName,
  };

  if (appearance === "primary") {
    return <PrimaryButton {...commonProps}>{children}</PrimaryButton>;
  }

  return <DefaultButton {...commonProps}>{children}</DefaultButton>;
};

const MessageBar: React.FC<{
  intent?: "error" | "warning" | "success" | "info";
  children?: React.ReactNode;
  className?: string;
}> = ({ intent = "info", children, className }) => (
  <FluentMessageBar
    className={className}
    messageBarType={
      intent === "error"
        ? MessageBarType.error
        : intent === "warning"
          ? MessageBarType.warning
          : intent === "success"
            ? MessageBarType.success
            : MessageBarType.info
    }
    isMultiline
  >
    {children}
  </FluentMessageBar>
);

const Badge: React.FC<{
  children?: React.ReactNode;
  appearance?: string;
  color?: string;
}> = ({ children }) => (
  <span className={styles.f2Badge}>
    {children}
  </span>
);

const Checkbox: React.FC<{
  checked?: boolean;
  onChange?: (event: React.FormEvent<HTMLElement> | undefined, data: { checked?: boolean }) => void;
  disabled?: boolean;
  label?: React.ReactNode;
}> = ({ checked, onChange, disabled, label }) => (
  <FluentCheckbox
    checked={checked}
    disabled={disabled}
    label={typeof label === "string" ? label : undefined}
    onRenderLabel={typeof label === "string" || typeof label === "undefined" ? undefined : () => <>{label}</>}
    onChange={(event, isChecked) =>
      onChange?.(event as React.FormEvent<HTMLElement>, { checked: isChecked })
    }
  />
);

const Dialog: React.FC<{ open: boolean; children?: React.ReactNode }> = ({ open, children }) => (
  <FluentDialog hidden={!open} modalProps={{ isBlocking: false }}>
    {children}
  </FluentDialog>
);

const DialogTrigger: React.FC<{ children?: React.ReactNode; disableButtonEnhancement?: boolean }> = ({ children }) => <>{children}</>;
const DialogSurface: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div>{children}</div>;
const DialogBody: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div>{children}</div>;
const DialogTitle: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div className={styles.f2DialogTitle}>{children}</div>
);
const DialogContent: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div>{children}</div>;
const DialogActions: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div className={styles.f2DialogActions}>{children}</div>
);
const Field: React.FC<{ children?: React.ReactNode }> = ({ children }) => <div>{children}</div>;

const ListManagement: React.FC<IListManagementProps> = ({
  siteContextService,
  onListCreated,
  className
}) => {
  const [sites, setSites] = React.useState<ISiteInfo[]>([]);
  const [listStatuses, setListStatuses] = React.useState<{ [siteId: string]: IAlertListStatus }>({});
  const [creatingList, setCreatingList] = React.useState<string | null>(null);
  const [message, setMessage] = React.useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [selectedLanguages, setSelectedLanguages] = React.useState<string[]>(['en-us']); // Default to English
  const [languageDialogOpen, setLanguageDialogOpen] = React.useState<{ siteId: string; siteName: string } | null>(null);

  const { loading, execute: loadSiteInformation } = useAsyncOperation(
    async () => {
      const siteHierarchy = siteContextService.getSitesHierarchy();
      setSites(siteHierarchy);
      const statuses: { [siteId: string]: IAlertListStatus } = {};
      for (const site of siteHierarchy) {
        try {
          statuses[site.id] = await siteContextService.getAlertListStatus(site);
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
      return statuses;
    },
    {
      onError: () => {
        setMessage({
          type: 'error',
          text: strings.FailedToLoadSiteInformation || 'Failed to load site information'
        });
      },
      logErrors: true
    }
  );

  React.useEffect(() => {
    loadSiteInformation();
  }, [loadSiteInformation]);

  const handleLanguageToggle = (languageCode: string) => {
    setSelectedLanguages(prev => {
      if (prev.includes(languageCode)) {
        if (languageCode === 'en-us') {
          return prev;
        }
        return prev.filter(code => code !== languageCode);
      } else {
        return [...prev, languageCode];
      }
    });
  };

  const handleOpenLanguageDialog = async (siteId: string, siteName: string) => {
    try {
      const status = listStatuses[siteId];
      if (status?.exists && status?.canAccess) {
        const configuredLanguages = await siteContextService.getSupportedLanguagesForSite(siteId);
        setSelectedLanguages(configuredLanguages);
      } else {
        setSelectedLanguages(['en-us']);
      }
      
      setLanguageDialogOpen({ siteId, siteName });
    } catch (error) {
      setSelectedLanguages(['en-us']);
      setLanguageDialogOpen({ siteId, siteName });
    }
  };

  const handleUpdateLanguages = async (siteId: string, siteName: string) => {
    try {
      setCreatingList(siteId);
      setMessage(null);
      setLanguageDialogOpen(null);

      const { SharePointAlertService } = await import('../Services/SharePointAlertService');
      const alertService = new SharePointAlertService(
        await siteContextService.getGraphClient(),
        siteContextService.getContext()
      );

      try {
        // Update supported languages
        await alertService.updateSupportedLanguages(siteId, selectedLanguages);
        
        setMessage({
          type: 'success',
          text: formatWithFallback(strings.LanguagesUpdatedSuccessfully, 'Languages updated successfully for {0}', siteName)
        });

        await siteContextService.refresh();
        await loadSiteInformation();
      } finally {
        // Context restoration no longer needed
      }
    } catch (error) {
      setMessage({
        type: 'error',
        text: error.message || formatWithFallback(strings.FailedToUpdateLanguages, 'Failed to update languages for {0}', siteName)
      });
    } finally {
      setCreatingList(null);
    }
  };

  const { execute: createListOperation } = useAsyncOperation(
    async (siteId: string, siteName: string) => {
      setCreatingList(siteId);
      setLanguageDialogOpen(null);

      const success = await siteContextService.createAlertsList(siteId, selectedLanguages);

      if (!success) {
        throw new Error('Creation failed');
      }

      const languagesList = selectedLanguages.length > 1
        ? ` with support for ${selectedLanguages.length} languages`
        : '';

      await siteContextService.refresh();
      await loadSiteInformation();

      if (onListCreated) {
        onListCreated(siteId);
      }

      return { siteId, siteName, languagesList };
    },
    {
      onSuccess: ({ siteName, languagesList }) => {
        setMessage({
          type: 'success',
          text: formatWithFallback(strings.AlertsListCreatedSuccessfully, 'Alerts list created successfully on {0}', siteName) + languagesList
        });
        setCreatingList(null);
      },
      onError: (error: Error) => {
        let errorMessage = error.message || strings.FailedToCreateAlertsList || 'Failed to create alerts list';

        if (error.message?.includes('LIST_INCOMPLETE')) {
          errorMessage = `List created but some features may be limited. ${error.message}`;
        } else if (error.message?.includes('PERMISSION_DENIED')) {
          errorMessage = 'Cannot create list: Insufficient permissions. Contact your SharePoint administrator.';
        }

        setMessage({
          type: error.message?.includes('LIST_INCOMPLETE') ? 'success' : 'error',
          text: errorMessage
        });
        setCreatingList(null);
      },
      logErrors: true
    }
  );

  const handleCreateList = React.useCallback(async (siteId: string, siteName: string) => {
    setMessage(null);
    await createListOperation(siteId, siteName);
  }, [createListOperation]);

  const getSiteIcon = (siteType: string) => {
    switch (siteType) {
      case 'home': return <Home24Regular />;
      case 'hub': return <Building24Regular />;
      default: return <Globe24Regular />;
    }
  };

  const getSiteTypeLabel = (siteType: string) => {
    switch (siteType) {
      case 'home': return strings.HomeSite || 'Home Site';
      case 'hub': return strings.HubSite || 'Hub Site';
      default: return strings.CurrentSite || 'Current Site';
    }
  };

  const getSiteDescription = (siteType: string) => {
    switch (siteType) {
      case 'home': return strings.HomeSiteDescription || 'Alerts shown on all sites in the tenant';
      case 'hub': return strings.HubSiteDescription || 'Alerts shown on hub and connected sites';
      default: return strings.CurrentSiteDescription || 'Alerts shown only on this site';
    }
  };

  const getStatusIcon = (status: IAlertListStatus) => {
    if (status.exists && status.canAccess) {
      return <CheckmarkCircle24Filled className={`${styles.statusIcon} ${styles.statusIconSuccess}`} />;
    } else if (status.exists && !status.canAccess) {
      return <Warning24Filled className={`${styles.statusIcon} ${styles.statusIconWarning}`} />;
    } else if (!status.exists && status.canCreate) {
      return <Add24Regular className={`${styles.statusIcon} ${styles.statusIconNeutral}`} />;
    } else {
      return <ErrorCircle24Filled className={`${styles.statusIcon} ${styles.statusIconError}`} />;
    }
  };

  const getStatusText = (status: IAlertListStatus) => {
    if (status.exists && status.canAccess) {
      return strings.ListExistsAndAccessible || 'List exists and accessible';
    } else if (status.exists && !status.canAccess) {
      return strings.ListExistsNoAccess || 'List exists but no access';
    } else if (!status.exists && status.canCreate) {
      return strings.ListNotExistsCanCreate || 'List not found - can create';
    } else {
      return strings.ListNotExistsCannotCreate || 'List not found - cannot create';
    }
  };

  if (loading) {
    return (
      <div className={`${styles.listManagement} ${className || ''}`}>
        <Card>
          <CardHeader
            image={<List24Regular />}
            header={<Text weight="semibold">{strings.AlertListsManagement || 'Alert Lists Management'}</Text>}
          />
          <CardPreview>
            <div className={styles.loadingContainer}>
              <Spinner label={strings.LoadingSiteInformation || 'Loading site information...'} />
            </div>
          </CardPreview>
        </Card>
      </div>
    );
  }

  return (
    <div className={`${styles.listManagement} ${className || ''}`}>
      {message && (
        <MessageBar intent={message.type} className={styles.messageBarWithMargin}>
          {message.text}
        </MessageBar>
      )}

      <Card>
        <CardHeader
          image={<List24Regular />}
          header={<Text weight="semibold">{strings.AlertListsManagement || 'Alert Lists Management'}</Text>}
          description={
            <Text size={200}>
              {strings.ManageAlertListsDescription || 'Manage alert lists across your site hierarchy'}
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
                      <Text size={200} className={styles.errorText}>
                        {status.error}
                      </Text>
                    )}
                  </div>

                  {!status.exists && status.canCreate && (
                    <div className={styles.createActions}>
                      <Dialog open={!!languageDialogOpen && languageDialogOpen.siteId === site.id}>
                        <DialogTrigger disableButtonEnhancement>
                          <Button
                            appearance="primary"
                            icon={<Add24Regular />}
                            onClick={() => handleOpenLanguageDialog(site.id, site.name)}
                            disabled={creatingList === site.id}
                          >
                            {creatingList === site.id 
                              ? (strings.CreatingList || 'Creating...')
                              : (strings.CreateAlertsList || 'Create Alerts List')
                            }
                          </Button>
                        </DialogTrigger>
                        <DialogSurface>
                          <DialogBody>
                            <DialogTitle>
                              <LocalLanguage24Regular className={styles.languageDialogIcon} />
                              {strings.SelectLanguagesForList || 'Select Languages for Alert List'}
                            </DialogTitle>
                            <DialogContent>
                              <Text>
                                {strings.SelectLanguagesDescription || 
                                  `Choose which languages to support for alerts on ${site.name}. English is required and will always be included.`
                                }
                              </Text>
                              
                              <div className={`${styles.languageGrid} ${styles.languageGridWithMargin}`}>
                                {AVAILABLE_LANGUAGES.map(language => (
                                  <Field key={language.code}>
                                    <Checkbox
                                      checked={selectedLanguages.includes(language.code)}
                                      onChange={() => handleLanguageToggle(language.code)}
                                      disabled={language.code === 'en-us'} // English is always required
                                      label={
                                        <div className={styles.languageLabel}>
                                          <Text weight="semibold">{language.nativeName}</Text>
                                          <Text size={200}>{language.name}</Text>
                                        </div>
                                      }
                                    />
                                  </Field>
                                ))}
                              </div>
                              
                              <div className={styles.languageSelectionSummary}>
                                <Text size={200}>
                                  <strong>{strings.SelectedLanguages || 'Selected languages'}:</strong> {selectedLanguages.length} 
                                  ({AVAILABLE_LANGUAGES
                                    .filter(lang => selectedLanguages.includes(lang.code))
                                    .map(lang => lang.nativeName)
                                    .join(', ')})
                                </Text>
                              </div>
                            </DialogContent>
                            <DialogActions>
                              <DialogTrigger disableButtonEnhancement>
                                <Button appearance="secondary" onClick={() => setLanguageDialogOpen(null)}>
                                  {strings.Cancel || 'Cancel'}
                                </Button>
                              </DialogTrigger>
                              <Button 
                                appearance="primary" 
                                onClick={() => {
                                  if (languageDialogOpen) {
                                    return handleCreateList(languageDialogOpen.siteId, languageDialogOpen.siteName);
                                  }
                                }}
                                disabled={creatingList === site.id}
                              >
                                {creatingList === site.id 
                                  ? (strings.CreatingList || 'Creating...')
                                  : formatWithFallback(strings.CreateWithSelectedLanguages, 'Create with {0} languages', selectedLanguages.length)
                                }
                             </Button>
                            </DialogActions>
                          </DialogBody>
                        </DialogSurface>
                      </Dialog>
                    </div>
                  )}

                  {status.exists && status.canAccess && (
                    <div className={styles.listInfo}>
                      <Text size={200} className={styles.successText}>
                        ✓ {strings.ReadyForAlerts || 'Ready for alerts'}
                      </Text>
                      <div className={styles.existingListActions}>
                        <Dialog open={!!languageDialogOpen && languageDialogOpen.siteId === site.id}>
                          <DialogTrigger disableButtonEnhancement>
                            <Button
                              appearance="subtle"
                              size="small"
                              icon={<LocalLanguage24Regular />}
                              onClick={() => handleOpenLanguageDialog(site.id, site.name)}
                            >
                              {strings.ViewEditLanguages || 'Languages'}
                            </Button>
                          </DialogTrigger>
                          <DialogSurface>
                            <DialogBody>
                              <DialogTitle>
                                <LocalLanguage24Regular className={styles.languageDialogIcon} />
                                {strings.ManageLanguagesForList || 'Manage Languages for Alert List'}
                              </DialogTitle>
                              <DialogContent>
                                <Text>
                                  {strings.ManageLanguagesDescription || 
                                    `Manage which languages are supported for alerts on ${site.name}. English is required and will always be included.`
                                  }
                                </Text>
                                
                                <div className={`${styles.languageGrid} ${styles.languageGridWithMargin}`}>
                                  {AVAILABLE_LANGUAGES.map(language => (
                                    <Field key={language.code}>
                                      <Checkbox
                                        checked={selectedLanguages.includes(language.code)}
                                        onChange={() => handleLanguageToggle(language.code)}
                                        disabled={language.code === 'en-us'} // English is always required
                                        label={
                                          <div className={styles.languageLabel}>
                                            <Text weight="semibold">{language.nativeName}</Text>
                                            <Text size={200}>{language.name}</Text>
                                          </div>
                                        }
                                      />
                                    </Field>
                                  ))}
                                </div>
                                
                                <div className={styles.languageSelectionSummary}>
                                  <Text size={200}>
                                    <strong>{strings.SelectedLanguages || 'Selected languages'}:</strong> {selectedLanguages.length} 
                                    ({AVAILABLE_LANGUAGES
                                      .filter(lang => selectedLanguages.includes(lang.code))
                                      .map(lang => lang.nativeName)
                                      .join(', ')})
                                  </Text>
                                </div>
                              </DialogContent>
                              <DialogActions>
                                <DialogTrigger disableButtonEnhancement>
                                  <Button appearance="secondary" onClick={() => setLanguageDialogOpen(null)}>
                                    {strings.Cancel || 'Cancel'}
                                  </Button>
                                </DialogTrigger>
                                <Button 
                                  appearance="primary" 
                                  onClick={() => {
                                    if (languageDialogOpen) {
                                      return handleUpdateLanguages(languageDialogOpen.siteId, languageDialogOpen.siteName);
                                    }
                                  }}
                                  disabled={creatingList === site.id}
                                >
                                  {strings.UpdateLanguages || 'Update Languages'}
                                </Button>
                              </DialogActions>
                            </DialogBody>
                          </DialogSurface>
                        </Dialog>
                      </div>
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
          header={<Text weight="semibold">{strings.AlertHierarchy || 'Alert Display Hierarchy'}</Text>}
        />
        <CardPreview>
          <div className={styles.hierarchyDescription}>
            <Text size={200}>
              {strings.AlertHierarchyDescription || 
                'Alerts are displayed based on site hierarchy: Home Site alerts appear everywhere, Hub Site alerts appear on hub and connected sites, and Site alerts appear only on the specific site.'}
            </Text>
            
            <div className={styles.hierarchyList}>
              <div className={styles.hierarchyItem}>
                <Home24Regular />
                <Text size={200}><strong>{strings.HomeSite || 'Home Site'}:</strong> {strings.HomeSiteScope || 'Shown on all sites'}</Text>
              </div>
              <div className={styles.hierarchyItem}>
                <Building24Regular />
                <Text size={200}><strong>{strings.HubSite || 'Hub Site'}:</strong> {strings.HubSiteScope || 'Shown on hub and connected sites'}</Text>
              </div>
              <div className={styles.hierarchyItem}>
                <Globe24Regular />
                <Text size={200}><strong>{strings.CurrentSite || 'Site'}:</strong> {strings.SiteScope || 'Shown only on this site'}</Text>
              </div>
            </div>
          </div>
        </CardPreview>
      </Card>
    </div>
  );
};

export default ListManagement;
