import * as React from "react";
import { logger } from '../Services/LoggerService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import {
  Search24Regular,
  CheckmarkCircle24Regular,
  Building24Regular,
  Globe24Regular,
  People24Regular,
  Home24Regular,
  Settings24Regular,
  Filter24Regular,
  ChevronDown24Regular,
  ChevronUp24Regular
} from "@fluentui/react-icons";
import { SharePointButton, SharePointInput, SharePointToggle } from "./SharePointControls";
import { ISiteOption, SiteContextDetector } from "../Utils/SiteContextDetector";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from "./SiteSelector.module.scss";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text as CoreText } from '@microsoft/sp-core-library';

export interface ISiteSelectorProps {
  selectedSites: string[];
  onSitesChange: (siteIds: string[]) => void;
  siteDetector: SiteContextDetector;
  graphClient: MSGraphClientV3;
  maxSelection?: number;
  allowMultiple?: boolean;
  showPermissionStatus?: boolean;
  className?: string;
}

interface IFilterOptions {
  showOnlyWritable: boolean;
  siteType: 'all' | 'hub' | 'team' | 'communication' | 'homesite';
  searchTerm: string;
}

const SiteSelector: React.FC<ISiteSelectorProps> = ({
  selectedSites,
  onSitesChange,
  siteDetector,
  graphClient,
  maxSelection,
  allowMultiple = true,
  showPermissionStatus = true,
  className
}) => {
  const [availableSites, setAvailableSites] = React.useState<ISiteOption[]>([]);
  const [searchTerm, setSearchTerm] = React.useState('');
  const [filters, setFilters] = React.useState<IFilterOptions>({
    showOnlyWritable: true,
    siteType: 'all',
    searchTerm: ''
  });
  const [showFilters, setShowFilters] = React.useState(false);
  const [suggestedSites, setSuggestedSites] = React.useState<{
    currentSite: ISiteOption;
    hubSites?: ISiteOption[];
    homesite?: ISiteOption;
    recentSites: ISiteOption[];
    followedSites: ISiteOption[];
  } | null>(null);

  // Load sites using useAsyncOperation
  const { loading, execute: loadSites } = useAsyncOperation(
    async () => {
      const [sites, suggestions] = await Promise.all([
        siteDetector.getAvailableSites(showPermissionStatus),
        siteDetector.getSuggestedDistributionScopes()
      ]);
      return { sites, suggestions };
    },
    {
      onSuccess: (result) => {
        if (result) {
          setAvailableSites(result.sites);
          setSuggestedSites(result.suggestions);
        }
      },
      onError: () => {
        logger.error('SiteSelector', 'Failed to load sites');
      },
      logErrors: true
    }
  );

  // Load sites on component mount
  React.useEffect(() => {
    loadSites();
  }, []);

  const filteredSites = React.useMemo(() => {
    let filtered = [...availableSites];

    // Apply search filter
    if (searchTerm.trim()) {
      const search = searchTerm.toLowerCase();
      filtered = filtered.filter(site =>
        site.name.toLowerCase().includes(search) ||
        site.url.toLowerCase().includes(search)
      );
    }

    // Apply permission filter
    if (filters.showOnlyWritable) {
      filtered = filtered.filter(site => site.userPermissions.canCreateAlerts);
    }

    // Apply site type filter
    if (filters.siteType !== 'all') {
      filtered = filtered.filter(site => site.type === filters.siteType);
    }

    // Sort by relevance: selected first, then by permission level, then alphabetically
    filtered.sort((a, b) => {
      const aSelected = selectedSites.includes(a.id);
      const bSelected = selectedSites.includes(b.id);

      if (aSelected && !bSelected) return -1;
      if (!aSelected && bSelected) return 1;

      const aCanWrite = a.userPermissions.canCreateAlerts;
      const bCanWrite = b.userPermissions.canCreateAlerts;

      if (aCanWrite && !bCanWrite) return -1;
      if (!aCanWrite && bCanWrite) return 1;

      return a.name.localeCompare(b.name);
    });

    return filtered;
  }, [availableSites, filters, searchTerm, selectedSites]);

  const toggleSiteSelection = React.useCallback((siteId: string) => {
    if (!allowMultiple) {
      onSitesChange([siteId]);
      return;
    }

    const isSelected = selectedSites.includes(siteId);
    let newSelection: string[];

    if (isSelected) {
      newSelection = selectedSites.filter(id => id !== siteId);
    } else {
      if (maxSelection && selectedSites.length >= maxSelection) {
        // Remove first selected and add new one
        newSelection = [...selectedSites.slice(1), siteId];
      } else {
        newSelection = [...selectedSites, siteId];
      }
    }

    onSitesChange(newSelection);
  }, [allowMultiple, selectedSites, maxSelection, onSitesChange]);

  const selectSuggestedScope = React.useCallback((scope: 'current' | 'hub' | 'homesite' | 'recent') => {
    if (!suggestedSites) return;

    let sitesToSelect: string[] = [];

    switch (scope) {
      case 'current':
        sitesToSelect = [suggestedSites.currentSite.id];
        break;
      case 'hub':
        if (suggestedSites.hubSites) {
          sitesToSelect = [
            suggestedSites.currentSite.id,
            ...suggestedSites.hubSites.map(s => s.id)
          ];
        }
        break;
      case 'homesite':
        if (suggestedSites.homesite) {
          sitesToSelect = [suggestedSites.homesite.id];
        }
        break;
      case 'recent':
        sitesToSelect = suggestedSites.recentSites
          .slice(0, 5)
          .filter(s => s.userPermissions.canCreateAlerts)
          .map(s => s.id);
        break;
    }

    onSitesChange(sitesToSelect);
  }, [suggestedSites, onSitesChange]);

  const getSiteIcon = React.useCallback((site: ISiteOption) => {
    switch (site.type) {
      case 'hub':
        return <Settings24Regular />;
      case 'homesite':
        return <Home24Regular />;
      case 'team':
        return <People24Regular />;
      case 'communication':
        return <Globe24Regular />;
      default:
        return <Building24Regular />;
    }
  }, []);

  const getPermissionBadge = React.useCallback((site: ISiteOption) => {
    if (!showPermissionStatus) return null;

    const { permissionLevel, canCreateAlerts } = site.userPermissions;

    if (!canCreateAlerts) {
      return <span className={`${styles.permissionBadge} ${styles.readOnly}`}>{strings.SiteSelectorPermissionReadOnly}</span>;
    }

    switch (permissionLevel) {
      case 'owner':
        return <span className={`${styles.permissionBadge} ${styles.owner}`}>{strings.SiteSelectorPermissionOwner}</span>;
      case 'fullControl':
        return <span className={`${styles.permissionBadge} ${styles.owner}`}>{strings.SiteSelectorPermissionFullControl}</span>;
      case 'contribute':
        return <span className={`${styles.permissionBadge} ${styles.contribute}`}>{strings.SiteSelectorPermissionCanEdit}</span>;
      case 'design':
        return <span className={`${styles.permissionBadge} ${styles.design}`}>{strings.SiteSelectorPermissionDesigner}</span>;
      default:
        return <span className={`${styles.permissionBadge} ${styles.readOnly}`}>{strings.SiteSelectorPermissionReadOnly}</span>;
    }
  }, [showPermissionStatus]);

  if (loading) {
    return (
      <div className={`${styles.siteSelector} ${className || ''}`}>
        <div className={styles.loading}>
          <div className={styles.loadingSpinner}></div>
          <p>{strings.SiteSelectorLoading}</p>
        </div>
      </div>
    );
  }

  return (
    <div className={`${styles.siteSelector} ${className || ''}`}>
      {/* Suggested Scopes */}
      {suggestedSites && (
        <div className={styles.suggestedScopes}>
          <h4>{strings.SiteSelectorQuickSelectionTitle}</h4>
          <div className={styles.scopeButtons}>
            <SharePointButton
              variant="secondary"
              onClick={() => selectSuggestedScope('current')}
              className={styles.scopeButton}
            >
              <Building24Regular /> {strings.SiteSelectorCurrentSiteOnly}
            </SharePointButton>

            {suggestedSites.hubSites && suggestedSites.hubSites.length > 0 && (
              <SharePointButton
                variant="secondary"
                onClick={() => selectSuggestedScope('hub')}
                className={styles.scopeButton}
              >
                <Settings24Regular /> {CoreText.format(strings.SiteSelectorHubScope, (suggestedSites.hubSites.length + 1).toString())}
              </SharePointButton>
            )}

            {suggestedSites.homesite && (
              <SharePointButton
                variant="secondary"
                onClick={() => selectSuggestedScope('homesite')}
                className={styles.scopeButton}
              >
                <Home24Regular /> {strings.SiteSelectorOrganizationHome}
              </SharePointButton>
            )}

            {suggestedSites.recentSites.length > 0 && (
              <SharePointButton
                variant="secondary"
                onClick={() => selectSuggestedScope('recent')}
                className={styles.scopeButton}
              >
                {CoreText.format(
                  strings.SiteSelectorRecentSites,
                  suggestedSites.recentSites.filter(s => s.userPermissions.canCreateAlerts).length.toString()
                )}
              </SharePointButton>
            )}
          </div>
        </div>
      )}

      {/* Search and Filters */}
      <div className={styles.searchAndFilters}>
        <div className={styles.searchBox}>
          <SharePointInput
            label=""
            value={searchTerm}
            onChange={setSearchTerm}
            placeholder={strings.SiteSelectorSearchPlaceholder}
          />
          <Search24Regular className={styles.searchIcon} />
        </div>

        <div className={styles.filterSection}>
          <SharePointButton
            variant="secondary"
            onClick={() => setShowFilters(!showFilters)}
            className={styles.filterToggle}
          >
            <Filter24Regular />
            {strings.SiteSelectorFiltersLabel}
            {showFilters ? <ChevronUp24Regular /> : <ChevronDown24Regular />}
          </SharePointButton>

          {showFilters && (
            <div className={styles.filterOptions}>
              <SharePointToggle
                label={strings.SiteSelectorShowWritable}
                checked={filters.showOnlyWritable}
                onChange={(checked) => setFilters(prev => ({ ...prev, showOnlyWritable: checked }))}
              />

              <div className={styles.typeFilter}>
                <label>{strings.SiteSelectorSiteTypeLabel}</label>
                <select
                  value={filters.siteType}
                  onChange={(e) => setFilters(prev => ({
                    ...prev,
                    siteType: e.target.value as any
                  }))}
                  className={styles.typeSelect}
                >
                  <option value="all">{strings.SiteSelectorSiteTypeAll}</option>
                  <option value="hub">{strings.SiteSelectorSiteTypeHub}</option>
                  <option value="team">{strings.SiteSelectorSiteTypeTeam}</option>
                  <option value="communication">{strings.SiteSelectorSiteTypeCommunication}</option>
                  <option value="homesite">{strings.SiteSelectorSiteTypeHome}</option>
                </select>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Selection Summary */}
      {selectedSites.length > 0 && (
        <div className={styles.selectionSummary}>
          <p>
            <strong>{selectedSites.length}</strong> {selectedSites.length === 1 ? strings.SiteSelectorSiteSingular : strings.SiteSelectorSitePlural} {strings.SiteSelectorSelected}
            {maxSelection && ` ${CoreText.format(strings.SiteSelectorMaxSelection, maxSelection.toString())}`}
          </p>
        </div>
      )}

      {/* Sites Grid */}
      <div className={styles.sitesGrid}>
        {filteredSites.map(site => {
          const isSelected = selectedSites.includes(site.id);
          const canSelect = !maxSelection || selectedSites.length < maxSelection || isSelected;

          return (
            <div
              key={site.id}
              className={`${styles.siteCard} ${isSelected ? styles.selected : ''} ${!canSelect ? styles.disabled : ''}`}
              onClick={() => canSelect && toggleSiteSelection(site.id)}
            >
              <div className={styles.siteIcon}>
                {getSiteIcon(site)}
              </div>

              <div className={styles.siteInfo}>
                <h4 className={styles.siteName}>{site.name}</h4>
                <p className={styles.siteUrl}>{new URL(site.url).hostname + new URL(site.url).pathname}</p>
                {getPermissionBadge(site)}
              </div>

              <div className={styles.selectionIndicator}>
                {isSelected && <CheckmarkCircle24Regular />}
              </div>
            </div>
          );
        })}
      </div>

      {filteredSites.length === 0 && (
        <div className={styles.noResults}>
          <Search24Regular className={styles.noResultsIcon} />
          <h4>{strings.SiteSelectorNoResultsTitle}</h4>
          <p>
            {strings.SiteSelectorNoResultsDescription}
            {filters.showOnlyWritable && ` ${strings.SiteSelectorWritableFilterHint}`}
          </p>
        </div>
      )}
    </div>
  );
};

export default SiteSelector;
