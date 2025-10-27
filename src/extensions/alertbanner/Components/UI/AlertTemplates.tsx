import * as React from "react";
import { logger } from '../Services/LoggerService';
import { useAsyncOperation } from '../Hooks/useAsyncOperation';
import {
  Settings24Regular,
  Warning24Regular,
  Sparkle24Regular,
  Document24Regular,
  Shield24Regular,
  Book24Regular,
  Trophy24Regular,
  ArrowSync24Regular,
  Clock24Regular,
  Folder24Regular,
  Megaphone24Regular,
  Search24Regular,
  Pin24Regular,
  Alert24Regular,
  Info24Regular
} from "@fluentui/react-icons";
import { AlertPriority, NotificationType } from "../Alerts/IAlerts";
import { SharePointAlertService, IAlertItem } from "../Services/SharePointAlertService";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import styles from "./AlertTemplates.module.scss";
import * as strings from 'AlertBannerApplicationCustomizerStrings';
import { Text } from '@microsoft/sp-core-library';

export interface IAlertTemplate {
  id: string;
  name: string;
  description: string;
  icon: React.ReactElement;
  category: "maintenance" | "announcement" | "emergency" | "update" | "celebration" | "info" | "interruption" | "training";
  template: {
    title: string;
    description: string;
    priority: AlertPriority;
    notificationType: NotificationType;
    isPinned: boolean;
    linkUrl?: string;
    linkDescription?: string;
  };
}

interface IAlertTemplatesProps {
  onSelectTemplate: (template: IAlertTemplate) => void;
  graphClient: MSGraphClientV3;
  context: ApplicationCustomizerContext;
  alertService: SharePointAlertService;
  className?: string;
}

// Helper function to get icon for alert type
const getIconForAlertType = (alertType: string): React.ReactElement => {
  const type = alertType.toLowerCase();
  switch (type) {
    case 'maintenance': return <Settings24Regular />;
    case 'warning': return <Warning24Regular />;
    case 'info': return <Info24Regular />;
    case 'interruption': return <Warning24Regular />;
    default: return <Alert24Regular />;
  }
};

// Helper function to get category for alert type
const getCategoryForAlertType = (alertType: string): "maintenance" | "announcement" | "emergency" | "update" | "celebration" | "info" | "interruption" | "training" => {
  const type = alertType.toLowerCase();
  switch (type) {
    case 'maintenance': return 'maintenance';
    case 'warning': return 'emergency';
    case 'info': return 'info';
    case 'interruption': return 'interruption';
    default: return 'announcement';
  }
};

// Helper function to convert SharePoint alert to template
const convertAlertToTemplate = (alert: IAlertItem): IAlertTemplate => {
  return {
    id: alert.id,
    name: alert.title,
    description: Text.format(strings.AlertTemplatesTemplateDescription, alert.title),
    icon: getIconForAlertType(alert.AlertType),
    category: getCategoryForAlertType(alert.AlertType),
    template: {
      title: alert.title,
      description: alert.description,
      priority: alert.priority,
      notificationType: alert.notificationType as NotificationType,
      isPinned: alert.isPinned,
      linkUrl: alert.linkUrl || "",
      linkDescription: alert.linkDescription || strings.AlertTemplatesDefaultLinkDescription
    }
  };
};

const AlertTemplates: React.FC<IAlertTemplatesProps> = ({
  onSelectTemplate,
  graphClient,
  context,
  alertService,
  className
}) => {
  const [selectedCategory, setSelectedCategory] = React.useState<string>("all");
  const [searchTerm, setSearchTerm] = React.useState("");
  const [templates, setTemplates] = React.useState<IAlertTemplate[]>([]);

  const categories = [
    { id: "all", name: strings.AlertTemplatesCategoryAll, icon: <Folder24Regular /> },
    { id: "maintenance", name: strings.AlertTemplatesCategoryMaintenance, icon: <Settings24Regular /> },
    { id: "info", name: strings.AlertTemplatesCategoryInformation, icon: <Info24Regular /> },
    { id: "emergency", name: strings.AlertTemplatesCategoryEmergency, icon: <Warning24Regular /> },
    { id: "interruption", name: strings.AlertTemplatesCategoryInterruption, icon: <Warning24Regular /> },
    { id: "training", name: strings.AlertTemplatesCategoryTraining, icon: <Book24Regular /> },
    { id: "announcement", name: strings.AlertTemplatesCategoryAnnouncements, icon: <Megaphone24Regular /> }
  ];

  // Load templates from SharePoint using useAsyncOperation
  const { loading, execute: loadTemplates } = useAsyncOperation(
    async () => {
      const currentSiteId = context.pageContext.site.id.toString();
      const templateAlerts = await alertService.getTemplateAlerts(currentSiteId);
      const convertedTemplates = templateAlerts.map(convertAlertToTemplate);
      return convertedTemplates;
    },
    {
      onSuccess: (convertedTemplates) => setTemplates(convertedTemplates || []),
      onError: () => {
        logger.warn('AlertTemplates', 'Failed to load templates from SharePoint');
        setTemplates([]);
      },
      logErrors: true
    }
  );

  // Load templates on component mount
  React.useEffect(() => {
    loadTemplates();
  }, [alertService, context]);

  const filteredTemplates = templates.filter(template => {
    const matchesCategory = selectedCategory === "all" || template.category === selectedCategory;
    const matchesSearch = template.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      template.description.toLowerCase().includes(searchTerm.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  const handleTemplateSelect = (template: IAlertTemplate) => {
    onSelectTemplate(template);
  };

  return (
    <div className={`${styles.templatesContainer} ${className || ''}`}>
      <div className={styles.templatesHeader}>
        <h3>{strings.AlertTemplatesHeaderTitle}</h3>
        <p>{strings.AlertTemplatesHeaderDescription}</p>
      </div>

      <div className={styles.searchAndFilter}>
        <div className={styles.searchBox}>
          <input
            type="text"
            placeholder={strings.AlertTemplatesSearchPlaceholder}
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className={styles.searchInput}
          />
          <span className={styles.searchIcon}><Search24Regular /></span>
        </div>

        <div className={styles.categoryFilter}>
          {categories.map(category => (
            <button
              key={category.id}
              className={`${styles.categoryButton} ${selectedCategory === category.id ? styles.active : ''}`}
              onClick={() => setSelectedCategory(category.id)}
            >
              <span className={styles.categoryIcon}>{category.icon}</span>
              {category.name}
            </button>
          ))}
        </div>
      </div>

      {loading ? (
        <div className={styles.loadingContainer}>
          <div className={styles.loadingIcon}>📋</div>
          <h4>{strings.AlertTemplatesLoadingTitle}</h4>
          <p>{strings.AlertTemplatesLoadingDescription}</p>
        </div>
      ) : (
        <>
          <div className={styles.templatesGrid}>
            {filteredTemplates.map(template => (
              <div
                key={template.id}
                className={styles.templateCard}
                onClick={() => handleTemplateSelect(template)}
              >
                <div className={styles.templateIcon}>
                  {template.icon}
                </div>
                <div className={styles.templateContent}>
                  <h4 className={styles.templateName}>{template.name}</h4>
                  <p className={styles.templateDescription}>{template.description}</p>
                  <div className={styles.templateMeta}>
                  <span className={`${styles.priorityBadge} ${styles[template.template.priority]}`}>
                    {template.template.priority.toUpperCase()}
                  </span>
                  {template.template.isPinned && (
                    <span className={styles.pinnedBadge}><Pin24Regular /> {strings.AlertTemplatesPinnedBadge}</span>
                  )}
                  {template.template.notificationType !== NotificationType.None && (
                    <span className={styles.notificationBadge}><Alert24Regular /> {strings.AlertTemplatesNotifyBadge}</span>
                  )}
                </div>
              </div>
              <div className={styles.templateAction}>
                <button className={styles.useTemplateButton}>
                  {strings.AlertTemplatesUseTemplate}
                </button>
              </div>
            </div>
          ))}
        </div>

        {!loading && filteredTemplates.length === 0 && templates.length > 0 && (
          <div className={styles.noResults}>
            <div className={styles.noResultsIcon}><Search24Regular /></div>
            <h4>{strings.AlertTemplatesNoResultsTitle}</h4>
            <p>{strings.AlertTemplatesNoResultsDescription}</p>
          </div>
        )}

        {!loading && templates.length === 0 && (
          <div className={styles.noResults}>
            <div className={styles.noResultsIcon}>📋</div>
            <h4>{strings.AlertTemplatesEmptyTitle}</h4>
            <p>{strings.AlertTemplatesEmptyDescription}</p>
          </div>
        )}
        </>
      )}
    </div>
  );
};

export default AlertTemplates;
