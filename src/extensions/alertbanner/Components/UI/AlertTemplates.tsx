import * as React from "react";
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
  Alert24Regular
} from "@fluentui/react-icons";
import { AlertPriority, NotificationType } from "../Alerts/IAlerts";
import styles from "./AlertTemplates.module.scss";

export interface IAlertTemplate {
  id: string;
  name: string;
  description: string;
  icon: React.ReactElement;
  category: "maintenance" | "announcement" | "emergency" | "update" | "celebration";
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
  className?: string;
}

const ALERT_TEMPLATES: IAlertTemplate[] = [
  {
    id: "maintenance",
    name: "Scheduled Maintenance",
    description: "Notify users about upcoming system maintenance",
    icon: <Settings24Regular />,
    category: "maintenance",
    template: {
      title: "Scheduled System Maintenance",
      description: "<p>We will be performing scheduled maintenance on <strong>[DATE]</strong> from <strong>[START TIME]</strong> to <strong>[END TIME]</strong>.</p><p>During this time, some services may be temporarily unavailable. We apologize for any inconvenience.</p>",
      priority: AlertPriority.High,
      notificationType: NotificationType.Browser,
      isPinned: true,
      linkUrl: "",
      linkDescription: "View maintenance details"
    }
  },
  {
    id: "outage",
    name: "Service Outage",
    description: "Alert users about service disruptions",
    icon: <Warning24Regular />,
    category: "emergency",
    template: {
      title: "Service Disruption Notice",
      description: "<p><strong>We are currently experiencing technical difficulties</strong> with [SERVICE NAME].</p><p>Our technical team is working to resolve this issue as quickly as possible. We will provide updates as they become available.</p><p>We apologize for the inconvenience.</p>",
      priority: AlertPriority.Critical,
      notificationType: NotificationType.Both,
      isPinned: true,
      linkUrl: "",
      linkDescription: "Check status page"
    }
  },
  {
    id: "new-feature",
    name: "New Feature Announcement",
    description: "Introduce new features or updates to users",
    icon: <Sparkle24Regular />,
    category: "announcement",
    template: {
      title: "New Feature: [FEATURE NAME]",
      description: "<p>We're excited to announce <strong>[FEATURE NAME]</strong>!</p><p>[FEATURE DESCRIPTION]. This enhancement will help you [BENEFIT].</p><p>The feature is now available to all users.</p>",
      priority: AlertPriority.Medium,
      notificationType: NotificationType.None,
      isPinned: false,
      linkUrl: "",
      linkDescription: "Learn more about the new feature"
    }
  },
  {
    id: "policy-update",
    name: "Policy Update",
    description: "Communicate important policy or procedure changes",
    icon: <Document24Regular />,
    category: "announcement",
    template: {
      title: "Important Policy Update",
      description: "<p>We have updated our <strong>[POLICY NAME]</strong> effective <strong>[EFFECTIVE DATE]</strong>.</p><p>Key changes include:</p><ul><li>[CHANGE 1]</li><li>[CHANGE 2]</li><li>[CHANGE 3]</li></ul><p>Please review the updated policy and ensure compliance.</p>",
      priority: AlertPriority.High,
      notificationType: NotificationType.Browser,
      isPinned: true,
      linkUrl: "",
      linkDescription: "Read full policy"
    }
  },
  {
    id: "security-alert",
    name: "Security Alert",
    description: "Urgent security-related notifications",
    icon: <Shield24Regular />,
    category: "emergency",
    template: {
      title: "Security Alert: Action Required",
      description: "<p><strong>Important Security Notice</strong></p><p>We have detected [SECURITY ISSUE]. As a precautionary measure, please:</p><ul><li>[ACTION 1]</li><li>[ACTION 2]</li><li>[ACTION 3]</li></ul><p>If you have any concerns, please contact IT Security immediately.</p>",
      priority: AlertPriority.Critical,
      notificationType: NotificationType.Both,
      isPinned: true,
      linkUrl: "",
      linkDescription: "Contact IT Security"
    }
  },
  {
    id: "training",
    name: "Training Reminder",
    description: "Remind users about mandatory training",
    icon: <Book24Regular />,
    category: "announcement",
    template: {
      title: "Training Reminder: [TRAINING NAME]",
      description: "<p>This is a reminder that <strong>[TRAINING NAME]</strong> is due by <strong>[DUE DATE]</strong>.</p><p>This training is mandatory for all [AUDIENCE]. Please complete it as soon as possible to maintain compliance.</p><p>Estimated completion time: [DURATION]</p>",
      priority: AlertPriority.Medium,
      notificationType: NotificationType.Browser,
      isPinned: false,
      linkUrl: "",
      linkDescription: "Start training"
    }
  },
  {
    id: "celebration",
    name: "Celebration/Achievement",
    description: "Share good news and achievements",
    icon: <Trophy24Regular />,
    category: "celebration",
    template: {
      title: "Congratulations! [ACHIEVEMENT]",
      description: "<p>We're thrilled to announce <strong>[ACHIEVEMENT DETAILS]</strong>!</p><p>This success is thanks to the hard work and dedication of our entire team. [ADDITIONAL DETAILS]</p><p>Thank you to everyone who contributed to this milestone!</p>",
      priority: AlertPriority.Low,
      notificationType: NotificationType.None,
      isPinned: false,
      linkUrl: "",
      linkDescription: "Learn more"
    }
  },
  {
    id: "system-update",
    name: "System Update",
    description: "Inform users about system updates and improvements",
    icon: <ArrowSync24Regular />,
    category: "update",
    template: {
      title: "System Update: [VERSION]",
      description: "<p>We've updated our system to version <strong>[VERSION]</strong>.</p><p><strong>What's new:</strong></p><ul><li>[IMPROVEMENT 1]</li><li>[IMPROVEMENT 2]</li><li>[BUG FIX 1]</li></ul><p>The update has been automatically applied and no action is required from users.</p>",
      priority: AlertPriority.Low,
      notificationType: NotificationType.None,
      isPinned: false,
      linkUrl: "",
      linkDescription: "View release notes"
    }
  },
  {
    id: "deadline-reminder",
    name: "Deadline Reminder",
    description: "Remind users about important deadlines",
    icon: <Clock24Regular />,
    category: "announcement",
    template: {
      title: "Reminder: [TASK] Due [DATE]",
      description: "<p><strong>Friendly Reminder:</strong> [TASK DESCRIPTION] is due on <strong>[DUE DATE]</strong>.</p><p>Please ensure you complete this task by the deadline. If you need assistance or have questions, please don't hesitate to reach out.</p><p>Days remaining: <strong>[DAYS_LEFT]</strong></p>",
      priority: AlertPriority.Medium,
      notificationType: NotificationType.Browser,
      isPinned: false,
      linkUrl: "",
      linkDescription: "Complete task"
    }
  }
];

const AlertTemplates: React.FC<IAlertTemplatesProps> = ({
  onSelectTemplate,
  className
}) => {
  const [selectedCategory, setSelectedCategory] = React.useState<string>("all");
  const [searchTerm, setSearchTerm] = React.useState("");

  const categories = [
    { id: "all", name: "All Templates", icon: <Folder24Regular /> },
    { id: "maintenance", name: "Maintenance", icon: <Settings24Regular /> },
    { id: "announcement", name: "Announcements", icon: <Megaphone24Regular /> },
    { id: "emergency", name: "Emergency", icon: <Warning24Regular /> },
    { id: "update", name: "Updates", icon: <ArrowSync24Regular /> },
    { id: "celebration", name: "Celebrations", icon: <Trophy24Regular /> }
  ];

  const filteredTemplates = ALERT_TEMPLATES.filter(template => {
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
        <h3>Choose a Template</h3>
        <p>Start with a pre-configured template and customize it to your needs</p>
      </div>

      <div className={styles.searchAndFilter}>
        <div className={styles.searchBox}>
          <input
            type="text"
            placeholder="Search templates..."
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
                  <span className={styles.pinnedBadge}><Pin24Regular /> PINNED</span>
                )}
                {template.template.notificationType !== NotificationType.None && (
                  <span className={styles.notificationBadge}><Alert24Regular /> NOTIFY</span>
                )}
              </div>
            </div>
            <div className={styles.templateAction}>
              <button className={styles.useTemplateButton}>
                Use Template →
              </button>
            </div>
          </div>
        ))}
      </div>

      {filteredTemplates.length === 0 && (
        <div className={styles.noResults}>
          <div className={styles.noResultsIcon}><Search24Regular /></div>
          <h4>No templates found</h4>
          <p>Try adjusting your search terms or category filter</p>
        </div>
      )}
    </div>
  );
};

export default AlertTemplates;