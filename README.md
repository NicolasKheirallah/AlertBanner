# Alert Header SPFx Extension

## Summary

The Alert Header SPFx Extension is a custom SharePoint Framework (SPFx) extension that displays alert notifications in the header of Modern SharePoint sites. Alerts are retrieved from a SharePoint list using the Microsoft Graph API, providing users with important updates and information seamlessly integrated with Microsoft 365 services.

![screenshot](https://github.com/NicolasKheirallah/AlertHeader/blob/main/Screenshot/Screenshot2024-08-17170932.png)

This project is inspired by the work of Thomas Daly but has been significantly rewritten and updated to leverage the Microsoft Graph API for enhanced functionality. Special thanks to Thomas Daly for the original concept!

[Thomas Daly Blog](https://github.com/tom-daly/alerts-header)

## Goal of this project

As a alert banner is often requested by organization such as IT but isn't readily available, I wanted to create one that could be used by any organization!

I also rarely get to code these days so I wanted to freshen up my knowledge

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.19.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to the [Microsoft 365 developer program](http://aka.ms/o365devprogram).

## Prerequisites

- Node.js (v18.x or later)
- React 17
- SPFX (v1.19.x or later)
- A SharePoint Online site collection
- Appropriate permissions to access and configure the tenant App Catalog

## Solution

| Solution     | Author(s)                                         |
| ------------ | ------------------------------------------------- |
| alert-header | [Nicolas Kheirallah](https://twitter.com/yourhandle) |

## Version history

| Version | Date            | Comments                                        |
| ------- | --------------- | ----------------------------------------------- |
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
- **Display Alerts**: Show alerts prominently in the header of Modern SharePoint pages.
- **User Interaction Handling**: Allows users to dismiss alerts, with the option to prevent dismissed alerts from reappearing.
- **Performance Optimization**: Utilizes local storage for caching alerts, improving performance.

## TODO

- ~~**Multi-Site Support**: Extend support for Root, Local, and Hub sites. ~~(Done)
- **Enhanced Design**: Improve design aesthetics and CSS customization options.
- **Advanced Sorting**: Implement sorting of alerts based on priority and date.

## Managing Alerts

- **Global Alerts**: Deployed across all sites, fetching alerts from the root site where the extension is installed.
- **Local Alerts**: After the extension is added to a site collection, a new list titled "Alerts" is automatically created in the Site Contents. To create a new alert, simply add a new item to this list.

### Alert List Configuration:

- **Title**: The main heading of the alert.
- **Description**: Detailed message content.
- **AlertType**: The type of alert, such as Info, Warning, Maintenance, or Interruption.
- **Link**: Optional URL for additional information or related content.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples, and open-source controls for your Microsoft 365 development.

## Concepts Demonstrated

This extension showcases:

- Integration of the Microsoft Graph API within SPFx extensions.
- Customizing the header section of Modern SharePoint pages.
- Efficient state management and caching using local and session storage.
- Responsive design and handling of user interactions within SPFx extensions.

> Contributions and suggestions to enhance this solution are highly encouraged. Share your feedback via GitHub or through the [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) program.
