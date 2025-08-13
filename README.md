# Alert Banner SPFx Extension

## Summary

The **Alert Banner SPFx Extension** is a custom SharePoint Framework (SPFx) extension designed to display alert notifications prominently in the Banner of Modern SharePoint sites. These alerts are dynamically retrieved from a SharePoint list using the Microsoft Graph API, ensuring users receive important updates and information seamlessly integrated with Microsoft 365 services.

![screenshot](https://github.com/NicolasKheirallah/alertbanner/blob/main/Screenshot/Screenshot2024-08-17170932.png)

This project draws inspiration from the work of Thomas Daly. Special thanks to Thomas Daly for the original concept!

[Thomas Daly alert banner](https://github.com/tom-daly/alerts-banner)

## Goal of this Project

Alert banners are frequently requested by organizations such as IT departments but are not readily available out-of-the-box. This extension aims to provide a flexible and reusable alert system that any organization can deploy with ease.

Additionally, this project serves as an opportunity to refresh and enhance coding skills within the SPFx ecosystem.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.19.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to the [Microsoft 365 developer program](http://aka.ms/o365devprogram).

## Prerequisites

- Node.js (v18.x or later)
- React 17
- SPFx (v1.21.1 or later)
- A SharePoint Online site collection
- Appropriate permissions to access and configure the tenant App Catalog

## Solution

| Solution     | Author(s)                                         |
| ------------ | ------------------------------------------------- |
| alert-banner | [Nicolas Kheirallah](https://github.com/nicolasKheirallah) |

## Version History

| Version | Date            | Comments                                        |
| ------- | --------------- | ----------------------------------------------- |
| 2.0     | August 13, 2025 | Major release: Hierarchical alert system with Home/Hub/Site distribution, automatic list management, multi-language content creation, and enhanced site context awareness |
| 1.6     | August 13, 2025 | Added comprehensive multi-language support with 8 languages, automatic language detection, and localization framework |
| 1.5     | August 12, 2025 | Removed dialog, integrated content directly into banner, added "Read More" functionality |
| 1.4     | March 5, 2025  | Enhanced UI with modern dialog and improved link styling |
| 1.3     | March 3, 2025   | Added alert prioritization, user targeting, notifications, rich media support |
| 1.2     | October 11, 2024| Added dynamic alerttypes, added support for homesite, hubsite and local site |
| 1.1     | August 17, 2024 | Added caching and session management for alerts |
| 1.0     | July 15, 2024   | Initial release                                 |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository.
- Navigate to the solution folder.
- In the command line, run:
  - **npm install**
  - **./buildProject.cmd**
  - **Deploy to app catalog**

> Additional steps may be required depending on your environment configuration.

## Features

This SPFx extension offers the following capabilities:

- **Fetch Alerts**: Retrieves alerts from a designated SharePoint list using the Microsoft Graph API.
- **Display Alerts**: Show alerts prominently in the banner of Modern SharePoint pages, with detailed content displayed directly within the expanded banner.
- **Dynamic Alert Type Configuration**: Configure alert types dynamically using a JSON property, allowing easy customization and scalability.
- **User Interaction Handling**: Allows users to dismiss alerts, with the option to prevent dismissed alerts from reappearing.
- **Performance Optimization**: Utilizes local storage for caching alerts, improving performance.
- **Multi-Site Support**: Support for Root, Local, and Hub sites.
- **Alert Prioritization**: Categorize alerts as Low, Medium, High, or Critical with visual differentiation.
- **User Targeting**: Target alerts to specific departments, job titles, or SharePoint groups.
- **Notification System**: Send browser and email notifications for important alerts.
- **Rich Media Support**: Include images, videos, HTML, and markdown content in alerts.
- **Quick Actions**: Add interactive buttons to alerts for direct user engagement.
- **Read More Functionality**: Automatically truncates long alert descriptions and provides a "Read More" button to expand the full content directly within the banner.
- **Stylized Action Links**: Buttons-styled links for better visibility and user interaction.
- **Multi-Language Support**: Full internationalization support with automatic language detection, manual language switching, and multi-language content creation.
- **Multi-Language Content Editor**: Create alert content in multiple languages with intelligent fallback systems and custom language support.
- **Hierarchical Alert System**: Intelligent alert distribution based on SharePoint site hierarchy (Home Site → Hub Site → Current Site).
- **Automatic List Management**: Smart detection and one-click creation of alert lists across site hierarchy.
- **Site Context Awareness**: Automatic detection of Home Sites, Hub Sites, and current site relationships.

## Multi-Language Support

The Alert Banner extension now includes comprehensive multi-language support with the following features:

### Supported Languages

- **English (en-us)** - English (Default)
- **French (fr-fr)** - Français
- **German (de-de)** - Deutsch
- **Spanish (es-es)** - Español
- **Swedish (sv-se)** - Svenska
- **Finnish (fi-fi)** - Suomi
- **Danish (da-dk)** - Dansk
- **Norwegian (nb-no)** - Norsk bokmål

### Language Features

- **Automatic Detection**: Detects user's language from SharePoint context or browser settings
- **Manual Selection**: Users can manually change language through the language selector in settings
- **Persistent Preferences**: Language choice is stored in browser local storage
- **Fallback Support**: Falls back to English if requested language is not available
- **Date/Time Localization**: Automatically formats dates and times according to selected locale
- **String Interpolation**: Dynamic content with parameter substitution
- **RTL Ready**: Framework supports right-to-left languages (when needed)

### For Administrators

Access the language selector through the Alert Settings dialog (available when in edit mode). The system will automatically detect the user's preferred language based on their SharePoint/browser settings, but users can manually override this selection.

For detailed information about localization features, development guidelines, and adding new languages, see [LOCALIZATION.md](./LOCALIZATION.md).

## Managing Alerts

The Alert Banner extension uses an intelligent hierarchical system to display alerts based on SharePoint site relationships:

### Alert Hierarchy

- **Home Site Alerts**: Displayed on ALL sites across the tenant - perfect for organization-wide announcements
- **Hub Site Alerts**: Displayed on the hub site and all sites connected to that hub - ideal for department or division updates  
- **Current Site Alerts**: Displayed only on the specific site - best for site-specific notifications

### Automatic Site Detection

The extension automatically detects:
- **Your organization's Home Site** (if configured)
- **Hub Site connections** for the current site
- **Site relationships** and permissions
- **Existing alert lists** across the hierarchy

### Alert List Management

The extension provides intelligent list management:
- **Automatic Detection**: Checks if alert lists exist on relevant sites
- **One-Click Creation**: Create missing alert lists directly from settings
- **Permission Validation**: Shows what you can access and create
- **Visual Status Indicators**: Clear icons show list availability and status

## Creating Multi-Language Content

The Alert Banner extension now supports creating content in multiple languages, allowing organizations to communicate effectively with diverse audiences.

### Using the Multi-Language Content Editor

1. **Access the Editor**: When creating or editing alerts through the Settings interface, you'll see a multi-language content section
2. **Select Languages**: Use the language tabs to switch between different language versions of your content
3. **Add Content**: Fill in the Title, Description, and Link Description for each language
4. **Language Management**: Add custom languages using the "Add Language" button

### Content Creation Workflow

1. **Primary Language**: Start by creating content in your primary language (typically English)
2. **Add Translations**: Use the language tabs to switch to other languages and provide translations
3. **Content Validation**: The system shows a summary of which languages have content
4. **Publish**: Save the alert - users will automatically see content in their preferred language

### Custom Language Support

Beyond the 8 built-in languages, administrators can add custom languages:

1. **Add Language**: Click "Add Language" in the multi-language editor
2. **Language Details**: Provide language code (e.g., 'it-it'), name, and native name
3. **Quick Setup**: Choose from suggested languages or manually configure
4. **Automatic Columns**: The system automatically creates SharePoint list columns for the new language

### Language Fallback System

The extension uses an intelligent fallback system:
- **Primary**: Shows content in user's preferred language
- **Secondary**: Falls back to English if preferred language content is missing  
- **Tertiary**: Uses default Title/Description fields as final fallback

### Best Practices for Multi-Language Content

- **Consistent Messaging**: Ensure translations convey the same meaning and urgency
- **Cultural Adaptation**: Consider cultural context, not just literal translation
- **Testing**: Preview alerts in different languages before publishing
- **Maintenance**: Keep translations updated when primary content changes

## Hierarchical Alert System

The Alert Banner extension implements a sophisticated three-tier alert system that follows SharePoint's natural site hierarchy:

### Alert Distribution Strategy

| Alert Level | Scope | Use Cases | Creation Location |
|-------------|--------|-----------|-------------------|
| **Home Site** | All sites in tenant | Organization announcements, company-wide news, emergency notifications | Home Site alerts list |
| **Hub Site** | Hub and connected sites | Department updates, division-wide communications, shared resources | Hub Site alerts list |
| **Site-Specific** | Individual site only | Team notifications, project updates, site maintenance | Current site alerts list |

### How It Works

1. **Automatic Detection**: The extension scans your SharePoint environment to identify:
   - Home Site (if configured at tenant level)
   - Hub Site connections for the current site
   - Current site context and permissions

2. **Intelligent Querying**: Only queries sites that:
   - Have alert lists created
   - User has permission to access
   - Are relevant to the current site context

3. **Hierarchical Display**: Alerts are displayed in priority order:
   - Home Site alerts (highest priority - shown everywhere)
   - Hub Site alerts (medium priority - shown on hub and connected sites)
   - Site alerts (local priority - shown on specific site only)

### List Management Interface

The Settings dialog includes a comprehensive List Management section that provides:

- **Site Hierarchy Visualization**: See your Home Site, Hub Site, and Current Site relationships
- **List Status Indicators**: Visual status for each site's alert list (✅ Available, ⚠️ Limited Access, ❌ Missing, ➕ Can Create)
- **One-Click List Creation**: Create missing alert lists with full field schema
- **Permission Validation**: Clear indication of what actions are available
- **Real-Time Updates**: Status updates immediately when lists are created

### Best Practices

#### For Home Site Alerts
- Use for critical, tenant-wide communications
- Keep content concise and universally relevant
- Consider multi-language content for global organizations
- Use high priority levels sparingly

#### For Hub Site Alerts  
- Focus on department or division-specific content
- Leverage user targeting for role-specific messages
- Coordinate with hub site owners for content strategy
- Use medium priority for most communications

#### For Site-Specific Alerts
- Perfect for team notifications and project updates
- Use lower priority levels for routine communications
- Leverage rich media for detailed explanations
- Target specific user groups as needed

### Alert List Fields & Configuration:

When creating the "Alerts" list in SharePoint, you should configure the following fields:

| Field Name      | Field Type                     | Description                                        |
|-----------------|--------------------------------|----------------------------------------------------|
| Title           | Single line of text            | The main heading of the alert                      |
| Description     | Multiple lines of text (Rich)  | Detailed message content                           |
| AlertType       | Choice                         | Type of alert (Info, Warning, etc.)                |
| Priority        | Choice                         | Low, Medium, High, Critical                        |
| IsPinned        | Yes/No                         | Whether the alert should be pinned to the top      |
| StartDateTime   | Date and Time                  | When the alert should begin displaying             |
| EndDateTime     | Date and Time                  | When the alert should stop displaying              |
| TargetingRules  | Multiple lines of text (Plain) | JSON defining which users should see the alert     |
| NotificationType| Choice                         | None, Browser, Email, Both                         |
| RichMedia       | Multiple lines of text (Plain) | JSON defining embedded media content               |
| QuickActions    | Multiple lines of text (Plain) | JSON defining interactive buttons for the alert    |
| Link            | Hyperlink or Picture           | URL for additional information                     |

#### Multi-Language Content Fields:

**Important:** When setting up the SharePoint list, include these additional fields for multi-language content support:

| Field Name      | Field Type                     | Description                                        |
|-----------------|--------------------------------|----------------------------------------------------|
| Title_EN        | Single line of text            | English version of the alert title                |
| Title_FR        | Single line of text            | French version of the alert title                 |
| Title_DE        | Single line of text            | German version of the alert title                 |
| Title_ES        | Single line of text            | Spanish version of the alert title                |
| Title_SV        | Single line of text            | Swedish version of the alert title                |
| Title_FI        | Single line of text            | Finnish version of the alert title                |
| Title_DA        | Single line of text            | Danish version of the alert title                 |
| Title_NO        | Single line of text            | Norwegian version of the alert title              |
| Description_EN  | Multiple lines of text (Rich)  | English version of the alert description          |
| Description_FR  | Multiple lines of text (Rich)  | French version of the alert description           |
| Description_DE  | Multiple lines of text (Rich)  | German version of the alert description           |
| Description_ES  | Multiple lines of text (Rich)  | Spanish version of the alert description          |
| Description_SV  | Multiple lines of text (Rich)  | Swedish version of the alert description          |
| Description_FI  | Multiple lines of text (Rich)  | Finnish version of the alert description          |
| Description_DA  | Multiple lines of text (Rich)  | Danish version of the alert description           |
| Description_NO  | Multiple lines of text (Rich)  | Norwegian version of the alert description        |
| LinkDescription_EN | Single line of text         | English version of the link description           |
| LinkDescription_FR | Single line of text         | French version of the link description            |
| LinkDescription_DE | Single line of text         | German version of the link description            |
| LinkDescription_ES | Single line of text         | Spanish version of the link description           |
| LinkDescription_SV | Single line of text         | Swedish version of the link description           |
| LinkDescription_FI | Single line of text         | Finnish version of the link description           |
| LinkDescription_DA | Single line of text         | Danish version of the link description            |
| LinkDescription_NO | Single line of text         | Norwegian version of the link description         |

**Note:** The extension automatically creates these fields when deployed. For additional languages, use the Language Management interface to add custom language columns dynamically.

#### Choice Field Options:

- **AlertType**: Create choice values matching your alert type configurations (Info, Warning, Maintenance, Interruption, etc.)
- **Priority**: Create these choice values: Low, Medium, High, Critical
- **NotificationType**: Create these choice values: none, browser, email, both

### User Interface Components

#### Alert Card
- **Clickable Interface**: The entire alert card is clickable, expanding and collapsing the banner to show/hide detailed content.
- **Visual Priority Indicators**: Color-coded borders and backgrounds based on alert priority.
- **Collapsible Content**: Toggle between condensed and expanded views.
- **Quick Action Buttons**: Customizable action buttons for direct user engagement.



#### Action Links
- **Icon Integration**: Link icon for visual clarity.
- **Button-Like Appearance**: Enhanced clickability compared to traditional text links.
- **Responsive Design**: Adapts to different screen sizes.

### Dynamic Alert Type Configuration:

Alert types can be customized dynamically using a JSON configuration property. This allows administrators to add, modify, or remove alert types without altering the codebase. The JSON structure defines the appearance and behavior of each alert type, including icons, colors, and additional styles.

**Example JSON Structure:**

```json
[
    {
       "name":"Info",
       "iconName":"Info12",
       "backgroundColor":"#389899",
       "textColor":"#ffffff",
       "additionalStyles":"",
       "priorityStyles": {
          "critical": "border: 2px solid #E81123;",
          "high": "border: 1px solid #EA4300;",
          "medium": "",
          "low": ""
       }
    },
    {
       "name":"Warning",
       "iconName":"ShieldAlert",
       "backgroundColor":"#f1c40f",
       "textColor":"#ffffff",
       "additionalStyles":""
    },
    {
       "name":"Maintenance",
       "iconName":"CRMServices",
       "backgroundColor":"#afd6d6",
       "textColor":"#ffffff",
       "additionalStyles":""
    },
    {
       "name":"Interruption",
       "iconName":"IncidentTriangle",
       "backgroundColor":"#c54644",
       "textColor":"#ffffff",
       "additionalStyles":""
    }
 ]
```

### User Targeting Configuration:

Alerts can be targeted to specific users based on their properties. This allows administrators to ensure that alerts are only shown to relevant audiences.

**Example Targeting Rules JSON (for TargetingRules field):**

```json
{
  "audiences": ["HR", "Finance"],
  "operation": "anyOf"
}
```

Operations supported:
- `anyOf`: User must match any of the audiences
- `allOf`: User must match all of the audiences
- `noneOf`: User must not match any of the audiences

### Rich Media Configuration:

Alerts can include various types of rich media content for enhanced communication.

**Example Rich Media JSON (for RichMedia field):**

```json
{
  "type": "markdown",
  "content": "# Important Update\n\nPlease review the [latest guidelines](https://example.com)",
  "altText": "Important update about new guidelines"
}
```

Media types supported:
- `image`: Display an image with the alert
- `video`: Embed a video player
- `html`: Include formatted HTML content
- `markdown`: Include markdown-formatted content

### Quick Actions Configuration:

Alerts can include interactive buttons for direct user engagement.

**Example Quick Actions JSON (for QuickActions field):**

```json
[
  {
    "label": "Add to Calendar",
    "actionType": "link",
    "url": "https://example.com/calendar",
    "icon": "Calendar"
  },
  {
    "label": "Acknowledge",
    "actionType": "acknowledge",
    "icon": "CheckMark"
  }
]
```

Action types supported:
- `link`: Open a URL
- `dismiss`: Dismiss the alert
- `acknowledge`: Acknowledge and dismiss the alert
- `custom`: Execute a custom JavaScript function

## Accessibility Features

The Alert Banner is designed with accessibility in mind:

- **Keyboard Navigation**: Full keyboard support for all interactive elements.
- **Screen Reader Support**: Proper ARIA attributes for better screen reader integration.
- **Focus Management**: Managed focus trap in dialogs to improve keyboard accessibility.
- **Color Contrast**: Meets WCAG 2.1 AA standards for color contrast.
- **Text Scaling**: Supports browser text scaling for vision impairments.
- **RTL Language Support**: Right-to-left language layout support for global accessibility.

## Performance Considerations

The Alert Banner is optimized for performance:

- **Efficient Rendering**: Uses React memo and callback optimizations to reduce render cycles.
- **Caching Strategy**: Implements local and session storage for efficient alert management.

- **Component Splitting**: Modular design allows for better code splitting.
- **CSS Optimization**: Carefully organized SCSS with variables for minimal CSS output.

## Deployment Guide

### Prerequisites for Deployment

1. **SharePoint Administrator Access**: Required to deploy to the tenant App Catalog
2. **Site Collection Administrator**: Needed for site-specific configurations
3. **Development Environment**: Node.js, SPFx CLI, and Visual Studio Code (recommended)

### Step-by-Step Deployment

#### 1. Build and Package
```bash
# Clone the repository
git clone https://github.com/NicolasKheirallah/AlertBanner.git
cd AlertBanner

# Install dependencies
npm install

# Build the solution
npm run build

# Bundle and package for production
gulp bundle --ship
gulp package-solution --ship
```

#### 2. Deploy to App Catalog
1. Navigate to your SharePoint tenant App Catalog
2. Upload the `.sppkg` file from the `/sharepoint/solution/` folder
3. When prompted, check "Make this solution available to all sites in the organization"
4. Click "Deploy"

#### 3. Add Extension to Sites
The extension can be deployed in three ways:

**Option A: Tenant-wide Deployment (Recommended)**
1. Go to SharePoint Admin Center
2. Navigate to "Advanced" > "Extensions"
3. Add the Alert Banner extension with appropriate scoping

**Option B: Site Collection Feature**
1. Go to Site Settings > Site Collection Features
2. Activate "Alert Banner Application Customizer"

**Option C: Manual Property Configuration**
Add the following to the site's User Custom Actions via PowerShell/CLI for Microsoft 365:

```powershell
m365 spo customaction add --webUrl "https://yourtenant.sharepoint.com/sites/yoursite" --location "ClientSideExtension.ApplicationCustomizer" --name "AlertBannerApplicationCustomizer" --clientSideComponentId "12345678-1234-1234-1234-123456789012"
```

### Configuration and Usage

#### Initial Setup
1. **Automatic Site Detection**: The extension detects your site hierarchy (Home Site, Hub Site, Current Site)
2. **List Status Check**: Verifies which sites have alert lists and your permissions
3. **Settings Access**: Open the Settings dialog in edit mode to manage the extension
4. **List Creation**: Use the integrated List Management to create missing alert lists with one click

#### Alert List Configuration
Use the field configuration table in the [Managing Alerts](#managing-alerts) section to set up your SharePoint list fields correctly.

#### User Permissions
- **Read permissions**: Required for viewing alerts
- **Contribute permissions**: Needed for dismissing alerts  
- **Design/Full Control**: Required for configuring alert settings

## Troubleshooting

### Common Issues and Solutions

#### Extension Not Visible After Deployment
1. **Clear browser cache**: Hard refresh (Ctrl+F5) or clear SharePoint cache
2. **Check deployment status**: Verify the .sppkg file was successfully deployed in App Catalog
3. **Verify feature activation**: Ensure the site collection feature is activated
4. **Check permissions**: User needs at least read access to the site

#### Alerts Not Displaying
1. **Site Hierarchy**: Check if alerts exist at the appropriate level (Home/Hub/Current site)
2. **List Management**: Use Settings → List Management to verify alert lists exist
3. **List fields**: Ensure all required fields exist with correct types
4. **Alert dates**: Check StartDateTime and EndDateTime fields
5. **User targeting**: Verify TargetingRules allow current user to see alerts
6. **Hub connections**: Verify site is connected to expected hub (if applicable)
7. **Home site**: Confirm organization has a configured Home Site (for tenant-wide alerts)
8. **Browser console**: Check for JavaScript errors or network issues

#### Language/Localization Issues
1. **Missing translations**: Check browser console for warnings about missing string keys
2. **Language not switching**: Clear browser local storage for the site
3. **Date format issues**: Verify browser locale settings
4. **Text overflow**: Some languages require more space - check responsive layout

#### Performance Issues
1. **Large alert lists**: Consider archiving old alerts or implementing pagination
2. **Rich media content**: Optimize image sizes and video formats
3. **Network latency**: Check SharePoint service health and connectivity

### Diagnostic Steps

1. **Site Hierarchy Check**: Use Settings → List Management to view detected site hierarchy
2. **List Status Verification**: Check which sites have alert lists and permission status
3. **Enable debug mode**: Add `#debug=1` to the URL to see additional console logging
4. **Check Network tab**: Monitor API calls to SharePoint in browser developer tools
5. **Hub Site Verification**: Confirm expected hub site connections in Site Settings
6. **Home Site Configuration**: Verify Home Site is configured at tenant level
7. **Inspect local storage**: Look for cached data and language preferences
8. **Review SharePoint logs**: Check tenant admin center for any error reports

### Getting Help

- **GitHub Issues**: Report bugs at [https://github.com/NicolasKheirallah/AlertBanner/issues](https://github.com/NicolasKheirallah/AlertBanner/issues)
- **Documentation**: Refer to [LOCALIZATION.md](./LOCALIZATION.md) for language-specific help
- **Community**: SharePoint development community forums and Microsoft Tech Community

## Current Limitations

- Animations and transitions are limited for cross-browser compatibility
- Outlook web integration is in development for a future release
- Custom notification sounds are not currently supported

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples, and open-source controls for your Microsoft 365 development.
- [FluentUI React Components](https://developer.microsoft.com/en-us/fluentui#/controls/web)

## Concepts Demonstrated

This extension showcases:

- **Integration of the Microsoft Graph API** within SPFx extensions for efficient data retrieval.
- **Hierarchical Site Architecture**: Intelligent detection and utilization of SharePoint's Home Site, Hub Site, and site relationships.
- **Multi-Site Alert Management**: Sophisticated querying across site hierarchy with permission-aware access control.
- **Dynamic Configuration Management**: Utilizing JSON properties to configure alert types, enhancing flexibility and maintainability.
- **Automatic List Detection and Creation**: Smart discovery of existing alert lists with one-click creation of missing infrastructure.
- **Site Context Awareness**: Comprehensive service for detecting and managing SharePoint site relationships and permissions.
- **Direct content display within the banner** of Modern SharePoint pages to provide a consistent and visible alerting mechanism.
- **Efficient State Management and Caching** using local and session storage to optimize performance and reduce redundant data fetching.
- **Multi-Language Content Management**: Column-based approach for creating and managing content in multiple languages.
- **Responsive Design and User Interaction Handling** to ensure alerts are accessible and user-friendly across various devices and screen sizes.
- **Advanced User Targeting** based on user properties and group membership.
- **Notification Integration** with browser notification API and Microsoft Graph email functionality.
- **Rich Media Content Rendering** with proper sanitization and responsive design.
- **Modern FluentUI Implementation** with proper component hierarchy and styling patterns.
- **Accessibility-First Design** ensuring usability for all users regardless of ability.