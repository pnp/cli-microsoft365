# Release notes

## v0.5.0

### New commands

**SharePoint Online:**

- [spo sitescript add](../cmd/spo/sitescript/sitescript-add.md) - adds site script for use with site designs [#65](https://github.com/SharePoint/office365-cli/issues/65)
- [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md) - lists site script available for use with site designs [#66](https://github.com/SharePoint/office365-cli/issues/66)
- [spo sitescript get](../cmd/spo/sitescript/sitescript-get.md) - gets information about the specified site script [#67](https://github.com/SharePoint/office365-cli/issues/67)
- [spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md) - removes the specified site script [#68](https://github.com/SharePoint/office365-cli/issues/68)
- [spo sitescript set](../cmd/spo/sitescript/sitescript-set.md) - updates existing site script [#216](https://github.com/SharePoint/office365-cli/issues/216)
- [spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md) - adds site design for creating modern sites [#69](https://github.com/SharePoint/office365-cli/issues/69)
- [spo list get](../cmd/spo/list/list-get.md) - gets information about the specific list [#199](https://github.com/SharePoint/office365-cli/issues/199)

### Changes

- fixed issue with prompts in non-interactive mode [#142](https://github.com/SharePoint/office365-cli/issues/142)
- added information about the current user to status commands [#202](https://github.com/SharePoint/office365-cli/issues/202)
- fixed issue with completing input that doesn't match commands [#222](https://github.com/SharePoint/office365-cli/issues/222)

## [v0.4.0](https://github.com/SharePoint/office365-cli/releases/tag/v0.4.0)

### Breaking changes

- renamed the `spo cdn origin set` command to [spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md) [#184](https://github.com/SharePoint/office365-cli/issues/184)

### New commands

**SharePoint Online:**

- [spo customaction list](../cmd/spo/customaction/customaction-list.md) - lists user custom actions for site or site collection [#19](https://github.com/SharePoint/office365-cli/issues/19)
- [spo site get](../cmd/spo/site/site-get.md) - gets information about the specific site collection [#114](https://github.com/SharePoint/office365-cli/issues/114)
- [spo site list](../cmd/spo/site/site-list.md) - lists modern sites of the given type [#115](https://github.com/SharePoint/office365-cli/issues/115)
- [spo site add](../cmd/spo/site/site-add.md) - creates new modern site [#116](https://github.com/SharePoint/office365-cli/issues/116)
- [spo app remove](../cmd/spo/app/app-remove.md) - removes the specified app from the tenant app catalog [#9](https://github.com/SharePoint/office365-cli/issues/9)
- [spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md) - creates a site collection app catalog in the specified site [#63](https://github.com/SharePoint/office365-cli/issues/63)
- [spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md) - removes site collection scoped app catalog from site [#64](https://github.com/SharePoint/office365-cli/issues/64)
- [spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md) - lists pending permission requests [#152](https://github.com/SharePoint/office365-cli/issues/152)
- [spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md) - approves the specified permission request [#153](https://github.com/SharePoint/office365-cli/issues/153)
- [spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md) - denies the specified permission request [#154](https://github.com/SharePoint/office365-cli/issues/154)
- [spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md) - lists permissions granted to the service principal [#155](https://github.com/SharePoint/office365-cli/issues/155)
- [spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md) - revokes the specified set of permissions granted to the service principal [#155](https://github.com/SharePoint/office365-cli/issues/156)
- [spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md) - enable or disable the service principal [#157](https://github.com/SharePoint/office365-cli/issues/157)
- [spo customaction add](../cmd/spo/customaction/customaction-add.md) - adds a user custom action for site or site collection [#18](https://github.com/SharePoint/office365-cli/issues/18)
- [spo externaluser list](../cmd/spo/externaluser/externaluser-list.md) - lists external users in the tenant [#27](https://github.com/SharePoint/office365-cli/issues/27)

**Azure Active Directory Graph:**

- [aad connect](../cmd/aad/connect.md) - connects to the Azure Active Directory Graph [#160](https://github.com/SharePoint/office365-cli/issues/160)
- [aad disconnect](../cmd/aad/disconnect.md) - disconnects from Azure Active Directory Graph [#161](https://github.com/SharePoint/office365-cli/issues/161)
- [aad status](../cmd/aad/status.md) - shows Azure Active Directory Graph connection status [#162](https://github.com/SharePoint/office365-cli/issues/162)
- [aad sp get](../cmd/aad/sp/sp-get.md) - gets information about the specific service principal [#158](https://github.com/SharePoint/office365-cli/issues/158)
- [aad oauth2grant list](../cmd/aad/oauth2grant/oauth2grant-list.md) - lists OAuth2 permission grants for the specified service principal [#159](https://github.com/SharePoint/office365-cli/issues/159)
- [aad oauth2grant add](../cmd/aad/oauth2grant/oauth2grant-add.md) - grant the specified service principal OAuth2 permissions to the specified resource [#164](https://github.com/SharePoint/office365-cli/issues/164)
- [aad oauth2grant set](../cmd/aad/oauth2grant/oauth2grant-set.md) - update OAuth2 permissions for the service principal [#163](https://github.com/SharePoint/office365-cli/issues/163)
- [aad oauth2grant remove](../cmd/aad/oauth2grant/oauth2grant-remove.md) - remove specified service principal OAuth2 permissions [#165](https://github.com/SharePoint/office365-cli/issues/165)

### Changes

- added support for persisting connection [#46](https://github.com/SharePoint/office365-cli/issues/46)
- fixed authentication bug in `spo app install`, `spo app uninstall` and `spo app upgrade` commands when connected to the tenant admin site [#118](https://github.com/SharePoint/office365-cli/issues/118)
- fixed authentication bug in the `spo customaction get` command when connected to the tenant admin site [#113](https://github.com/SharePoint/office365-cli/issues/113)
- fixed bug in rendering help for commands when using the `--help` option [#104](https://github.com/SharePoint/office365-cli/issues/104)
- added detailed output to the `spo customaction get` command [#93](https://github.com/SharePoint/office365-cli/issues/93)
- improved collecting telemetry [#130](https://github.com/SharePoint/office365-cli/issues/130), [#131](https://github.com/SharePoint/office365-cli/issues/131), [#132](https://github.com/SharePoint/office365-cli/issues/132), [#133](https://github.com/SharePoint/office365-cli/issues/133)
- added support for the `skipFeatureDeployment` flag to the [spo app deploy](../cmd/spo/app/app-deploy.md) command [#134](https://github.com/SharePoint/office365-cli/issues/134)
- wrapped executing commands in `try..catch` [#109](https://github.com/SharePoint/office365-cli/issues/109)
- added serializing objects in log [#108](https://github.com/SharePoint/office365-cli/issues/108)
- added support for autocomplete in Zsh, Bash and Fish and Clink (cmder) on Windows [#141](https://github.com/SharePoint/office365-cli/issues/141), [#190](https://github.com/SharePoint/office365-cli/issues/190)

## [v0.3.0](https://github.com/SharePoint/office365-cli/releases/tag/v0.3.0)

### New commands

**SharePoint Online:**

- [spo customaction get](../cmd/spo/customaction/customaction-get.md) - gets information about the specific user custom action [#20](https://github.com/SharePoint/office365-cli/issues/20)

### Changes

- changed command output to silent [#47](https://github.com/SharePoint/office365-cli/issues/47)
- added user-agent string to all requests [#52](https://github.com/SharePoint/office365-cli/issues/52)
- refactored `spo cdn get` and `spo storageentity set` to use the `getRequestDigest` helper [#78](https://github.com/SharePoint/office365-cli/issues/78) and [#80](https://github.com/SharePoint/office365-cli/issues/80)
- added common handler for rejected OData promises [#59](https://github.com/SharePoint/office365-cli/issues/59)
- added Google Analytics code to documentation [#84](https://github.com/SharePoint/office365-cli/issues/84)
- added support for formatting command output as JSON [#48](https://github.com/SharePoint/office365-cli/issues/48)

## [v0.2.0](https://github.com/SharePoint/office365-cli/releases/tag/v0.2.0)

### New commands

**SharePoint Online:**

- [spo app add](../cmd/spo/app/app-add.md) - add an app to the specified SharePoint Online app catalog [#3](https://github.com/SharePoint/office365-cli/issues/3)
- [spo app deploy](../cmd/spo/app/app-deploy.md) - deploy the specified app in the tenant app catalog [#7](https://github.com/SharePoint/office365-cli/issues/7)
- [spo app get](../cmd/spo/app/app-get.md) - get information about the specific app from the tenant app catalog [#2](https://github.com/SharePoint/office365-cli/issues/2)
- [spo app install](../cmd/spo/app/app-install.md) - install an app from the tenant app catalog in the site [#4](https://github.com/SharePoint/office365-cli/issues/4)
- [spo app list](../cmd/spo/app/app-list.md) - list apps from the tenant app catalog [#1](https://github.com/SharePoint/office365-cli/issues/1)
- [spo app retract](../cmd/spo/app/app-retract.md) - retract the specified app from the tenant app catalog [#8](https://github.com/SharePoint/office365-cli/issues/8)
- [spo app uninstall](../cmd/spo/app/app-uninstall.md) - uninstall an app from the site [#5](https://github.com/SharePoint/office365-cli/issues/5)
- [spo app upgrade](../cmd/spo/app/app-upgrade.md) - upgrade app in the specified site [#6](https://github.com/SharePoint/office365-cli/issues/6)

## v0.1.1

### Changes

- Fixed bug in resolving command paths on Windows

## v0.1.0

Initial release.

### New commands

**SharePoint Online:**

- [spo cdn get](../cmd/spo/cdn/cdn-get.md) - get Office 365 CDN status
- [spo cdn origin list](../cmd/spo/cdn/cdn-origin-list.md) - list Office 365 CDN origins
- [spo cdn origin remove](../cmd/spo/cdn/cdn-origin-remove.md) - remove Office 365 CDN origin
- [spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md) - add Office 365 CDN origin
- [spo cdn policy list](../cmd/spo/cdn/cdn-policy-list.md) - list Office 365 CDN policies
- [spo cdn policy set](../cmd/spo/cdn/cdn-policy-set.md) - set Office 365 CDN policy
- [spo cdn set](../cmd/spo/cdn/cdn-set.md) - enable/disable Office 365 CDN
- [spo connect](../cmd/spo/connect.md) - connect to a SharePoint Online site
- [spo disconnect](../cmd/spo/disconnect.md) - disconnect from SharePoint
- [spo status](../cmd/spo/status.md) - show SharePoint Online connection status
- [spo storageentity get](../cmd/spo/storageentity/storageentity-get.md) - get value of a tenant property
- [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md) - list all tenant properties
- [spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md) - remove a tenant property
- [spo storageentity set](../cmd/spo/storageentity/storageentity-set.md) - set a tenant property