# Rencore Deploy SPFx Package

[![Rencore logo](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/rencore.png)](https://rencore.com)

Easily deploy a SharePoint Framework solution package to SharePoint Online directly from Visual Studio Code.

## Deploy the solution

Choose if you want to deploy the solution in the app catalog or to all sites in your tenant.

![The 'Deploy in the app catalog' option highlighted in the context menu of a .sppkg file in VSCode](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/deploy-sppkg-in-app-catalog.png)
![The 'Deploy to all SharePoint sites' option highlighted in the context menu of a .sppkg file in VSCode](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/deploy-sppkg-to-all-sites.png)

_The extension will automatically upload the selected package to the tenant app catalog. If the package already exists, it will be overwritten. Only solutions that support global deployment, can be deployed globally._

## Connect to your tenant

Connect to your SharePoint Online tenant, by specifying its URL.

![The 'Rencore Deploy SPFx Package' extension prompting for specifying the SharePoint URL](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/sharepoint-url-prompt.png)

_You can specify the URL permanently using the `rencoreSpfxDeploy.sharePointUrl` setting. When this setting is specified, you won't be prompted for the SharePoint URL._

![The 'rencoreSpfxDeploy.sharePointUrl' setting highlighted in the VSCode settings](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/sharepoint-url-setting.png)

## Keep your credentials to yourself

When connecting to your tenant, the extension uses OAuth and doesn't have access to your credentials. It can only access your tenant after you granted it permissions and signed in with your account.

![The 'Rencore Deploy SPFx Package' extension prompts for completing Azure AD OAuth flow](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/oauth-prompt.png)

_At any time you can revoke the extension's permissions through Azure AD applications settings._

## Check connection status

At any time, check the status of your connection to SharePoint.

![SharePoint connection information displayed in VSCode](https://github.com/rencoreab/vscode-spfx-deploy-package/raw/master/./assets/connection-status.png)

## Release Notes

### 1.1.0

Added support for deploying to site collection app catalogs

### 1.0.2

Changed activation of the extension in case .sppkg isn't available yet

### 1.0.1

Added handling error when app catalog can't be found

### 1.0.0

Initial release