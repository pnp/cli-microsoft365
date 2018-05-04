# Release notes

## [v1.2.0](https://github.com/SharePoint/office365-cli/releases/tag/v1.2.0)

### New commands

**SharePoint Online:**

- [spo file remove](../cmd/spo/file/file-remove.md) - removes the specified file [#287](https://github.com/SharePoint/office365-cli/issues/287)
- [spo hubsite data get](../cmd/spo/hubsite/hubsite-data-get.md) - gets hub site data for the specified site [#394](https://github.com/SharePoint/office365-cli/issues/394)
- [spo hubsite theme sync](../cmd/spo/hubsite/hubsite-theme-sync.md) - applies any theme updates from the parent hub site [#401](https://github.com/SharePoint/office365-cli/issues/401)
- [spo listitem add](../cmd/spo/listitem/listitem-add.md) - creates a list item in the specified list [#270](https://github.com/SharePoint/office365-cli/issues/270)
- [spo listitem remove](../cmd/spo/listitem/listitem-remove.md) - removes the specified list item [#272](https://github.com/SharePoint/office365-cli/issues/272)
- [spo page control get](../cmd/spo/page/page-control-get.md) - gets information about the specific control on a modern page [#414](https://github.com/SharePoint/office365-cli/issues/414)
- [spo page control list](../cmd/spo/page/page-control-list.md) - lists controls on the specific modern page [#413](https://github.com/SharePoint/office365-cli/issues/413)
- [spo page get](../cmd/spo/page/page-get.md) - gets information about the specific modern page [#360](https://github.com/SharePoint/office365-cli/issues/360)
- [spo propertybag set](../cmd/spo/propertybag/propertybag-set.md) - sets the value of the specified property in the property bag [#393](https://github.com/SharePoint/office365-cli/issues/393)
- [spo web clientsidewebpart list](../cmd/spo/web/web-clientsidewebpart-list.md) - lists available client-side web parts [#367](https://github.com/SharePoint/office365-cli/issues/367)

**Microsoft Graph:**

- [graph user get](../cmd/graph/user/user-get.md) - gets information about the specified user [#326](https://github.com/SharePoint/office365-cli/issues/326)
- [graph user list](../cmd/graph/user/user-list.md) - lists users matching specified criteria [#327](https://github.com/SharePoint/office365-cli/issues/327)

### Changes

- added support for authenticating using credentials solving [#388](https://github.com/SharePoint/office365-cli/issues/388)

## [v1.1.0](https://github.com/SharePoint/office365-cli/releases/tag/v1.1.0)

### New commands

**SharePoint Online:**

- [spo file get](../cmd/spo/file/file-get.md) - gets information about the specified file [#282](https://github.com/SharePoint/office365-cli/issues/282)
- [spo page add](../cmd/spo/page/page-add.md) - creates modern page [#361](https://github.com/SharePoint/office365-cli/issues/361)
- [spo page list](../cmd/spo/page/page-list.md) - lists all modern pages in the given site [#359](https://github.com/SharePoint/office365-cli/issues/359)
- [spo page set](../cmd/spo/page/page-set.md) - updates modern page properties [#362](https://github.com/SharePoint/office365-cli/issues/362)
- [spo propertybag remove](../cmd/spo/propertybag/propertybag-remove.md) - removes specified property from the property bag [#291](https://github.com/SharePoint/office365-cli/issues/291)
- [spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md) - applies a site design to an existing site collection [#339](https://github.com/SharePoint/office365-cli/issues/339)
- [spo theme get](../cmd/spo/theme/theme-get.md) - gets custom theme information [#349](https://github.com/SharePoint/office365-cli/issues/349)
- [spo theme list](../cmd/spo/theme/theme-list.md) - retrieves the list of custom themes [#332](https://github.com/SharePoint/office365-cli/issues/332)
- [spo theme remove](../cmd/spo/theme/theme-remove.md) - removes existing theme [#331](https://github.com/SharePoint/office365-cli/issues/331)
- [spo theme set](../cmd/spo/theme/theme-set.md) - add or update a theme [#330](https://github.com/SharePoint/office365-cli/issues/330), [#340](https://github.com/SharePoint/office365-cli/issues/340)
- [spo web get](../cmd/spo/web/web-get.md) - retrieve information about the specified site [#188](https://github.com/SharePoint/office365-cli/issues/188)

**Microsoft Graph:**

- [graph o365group remove](../cmd/graph/o365group/o365group-remove.md) - removes an Office 365 Group [#309](https://github.com/SharePoint/office365-cli/issues/309)
- [graph o365group restore](../cmd/graph/o365group/o365group-restore.md) - restores a deleted Office 365 Group [#346](https://github.com/SharePoint/office365-cli/issues/346)
- [graph siteclassification get](../cmd/graph/siteclassification/siteclassification-get.md) - gets site classification configuration [#303](https://github.com/SharePoint/office365-cli/issues/303)

**Azure Management Service:**

- [azmgmt connect](../cmd/azmgmt/connect.md) - connects to the Azure Management Service [#378](https://github.com/SharePoint/office365-cli/issues/378)
- [azmgmt disconnect](../cmd/azmgmt/disconnect.md) - disconnects from the Azure Management Service [#378](https://github.com/SharePoint/office365-cli/issues/378)
- [azmgmt status](../cmd/azmgmt/status.md) - shows Azure Management Service connection status [#378](https://github.com/SharePoint/office365-cli/issues/378)
- [azmgmt flow environment get](../cmd/azmgmt/flow/flow-environment-get.md) - gets information about the specified Microsoft Flow environment [#380](https://github.com/SharePoint/office365-cli/issues/380)
- [azmgmt flow environment list](../cmd/azmgmt/flow/flow-environment-list.md) - lists Microsoft Flow environments in the current tenant [#379](https://github.com/SharePoint/office365-cli/issues/379)
- [azmgmt flow get](../cmd/azmgmt/flow/flow-get.md) - gets information about the specified Microsoft Flow [#382](https://github.com/SharePoint/office365-cli/issues/382)
- [azmgmt flow list](../cmd/azmgmt/flow/flow-list.md) - lists Microsoft Flows in the given environment [#381](https://github.com/SharePoint/office365-cli/issues/381)

### Updated commands

**Microsoft Graph:**

- [graph o365group list](../cmd/graph/o365group/o365group-list.md) - added support for listing deleted Office 365 Groups [#347](https://github.com/SharePoint/office365-cli/issues/347)

### Changes

- fixed bug in retrieving Office 365 groups in immersive mode solving [#351](https://github.com/SharePoint/office365-cli/issues/351)

## [v1.0.0](https://github.com/SharePoint/office365-cli/releases/tag/v1.0.0)

### Breaking changes

- switched to a custom Azure AD application for communicating with Office 365. After installing this version you have to reconnect to Office 365

### New commands

**SharePoint Online:**

- [spo file list](../cmd/spo/file/file-list.md) - lists all available files in the specified folder and site [#281](https://github.com/SharePoint/office365-cli/issues/281)
- [spo list add](../cmd/spo/list/list-add.md) - creates list in the specified site [#204](https://github.com/SharePoint/office365-cli/issues/204)
- [spo list remove](../cmd/spo/list/list-remove.md) - removes the specified list [#206](https://github.com/SharePoint/office365-cli/issues/206)
- [spo list set](../cmd/spo/list/list-set.md) - updates the settings of the specified list [#205](https://github.com/SharePoint/office365-cli/issues/205)
- [spo customaction clear](../cmd/spo/customaction/customaction-clear.md) - deletes all custom actions in the collection [#231](https://github.com/SharePoint/office365-cli/issues/231)
- [spo propertybag get](../cmd/spo/propertybag/propertybag-get.md) - gets the value of the specified property from the property bag [#289](https://github.com/SharePoint/office365-cli/issues/289)
- [spo propertybag list](../cmd/spo/propertybag/propertybag-list.md) - gets property bag values [#288](https://github.com/SharePoint/office365-cli/issues/288)
- [spo site set](../cmd/spo/site/site-set.md) - updates properties of the specified site [#121](https://github.com/SharePoint/office365-cli/issues/121)
- [spo site classic add](../cmd/spo/site/site-classic-add.md) - creates new classic site [#123](https://github.com/SharePoint/office365-cli/issues/123)
- [spo site classic set](../cmd/spo/site/site-classic-set.md) - change classic site settings [#124](https://github.com/SharePoint/office365-cli/issues/124)
- [spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md) - updates a site design with new values [#251](https://github.com/SharePoint/office365-cli/issues/251)
- [spo tenant appcatalogurl get](../cmd/spo/tenant/tenant-appcatalogurl-get.md) - gets the URL of the tenant app catalog [#315](https://github.com/SharePoint/office365-cli/issues/315)
- [spo web add](../cmd/spo/web/web-add.md) - create new subsite [#189](https://github.com/SharePoint/office365-cli/issues/189)
- [spo web list](../cmd/spo/web/web-list.md) - lists subsites of the specified site [#187](https://github.com/SharePoint/office365-cli/issues/187)
- [spo web remove](../cmd/spo/web/web-remove.md) - delete specified subsite [#192](https://github.com/SharePoint/office365-cli/issues/192)

**Microsoft Graph:**

- [graph connect](../cmd/graph/connect.md) - connects to the Microsoft Graph [#10](https://github.com/SharePoint/office365-cli/issues/10)
- [graph disconnect](../cmd/graph/disconnect.md) - disconnects from the Microsoft Graph [#10](https://github.com/SharePoint/office365-cli/issues/10)
- [graph status](../cmd/graph/status.md) - shows Microsoft Graph connection status [#10](https://github.com/SharePoint/office365-cli/issues/10)
- [graph o365group add](../cmd/graph/o365group/o365group-add.md) - creates Office 365 Group [#308](https://github.com/SharePoint/office365-cli/issues/308)
- [graph o365group get](../cmd/graph/o365group/o365group-get.md) - gets information about the specified Office 365 Group [#306](https://github.com/SharePoint/office365-cli/issues/306)
- [graph o365group list](../cmd/graph/o365group/o365group-list.md) - lists Office 365 Groups in the current tenant [#305](https://github.com/SharePoint/office365-cli/issues/305)
- [graph o365group set](../cmd/graph/o365group/o365group-set.md) - updates Office 365 Group properties [#307](https://github.com/SharePoint/office365-cli/issues/307)

### Changes

- fixed bug in logging dates [#317](https://github.com/SharePoint/office365-cli/issues/317)
- fixed typo in the example of the [spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md) command [#338](https://github.com/SharePoint/office365-cli/issues/338)

## [v0.5.0](https://github.com/SharePoint/office365-cli/releases/tag/v0.5.0)

### Breaking changes

- changed the [spo site get](../cmd/spo/site/site-get.md) command to return SPSite properties [#293](https://github.com/SharePoint/office365-cli/issues/293)

### New commands

**SharePoint Online:**

- [spo sitescript add](../cmd/spo/sitescript/sitescript-add.md) - adds site script for use with site designs [#65](https://github.com/SharePoint/office365-cli/issues/65)
- [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md) - lists site script available for use with site designs [#66](https://github.com/SharePoint/office365-cli/issues/66)
- [spo sitescript get](../cmd/spo/sitescript/sitescript-get.md) - gets information about the specified site script [#67](https://github.com/SharePoint/office365-cli/issues/67)
- [spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md) - removes the specified site script [#68](https://github.com/SharePoint/office365-cli/issues/68)
- [spo sitescript set](../cmd/spo/sitescript/sitescript-set.md) - updates existing site script [#216](https://github.com/SharePoint/office365-cli/issues/216)
- [spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md) - adds site design for creating modern sites [#69](https://github.com/SharePoint/office365-cli/issues/69)
- [spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md) - gets information about the specified site design [#86](https://github.com/SharePoint/office365-cli/issues/86)
- [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md) - lists available site designs for creating modern sites [#85](https://github.com/SharePoint/office365-cli/issues/85)
- [spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md) - removes the specified site design [#87](https://github.com/SharePoint/office365-cli/issues/87)
- [spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md) - grants access to a site design for one or more principals [#88](https://github.com/SharePoint/office365-cli/issues/88)
- [spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md) - revokes access from a site design for one or more principals [#89](https://github.com/SharePoint/office365-cli/issues/89)
- [spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md) - gets a list of principals that have access to a site design [#90](https://github.com/SharePoint/office365-cli/issues/90)
- [spo list get](../cmd/spo/list/list-get.md) - gets information about the specific list [#199](https://github.com/SharePoint/office365-cli/issues/199)
- [spo customaction remove](../cmd/spo/customaction/customaction-remove.md) - removes the specified custom action [#21](https://github.com/SharePoint/office365-cli/issues/21)
- [spo site classic list](../cmd/spo/site/site-classic-list.md) - lists sites of the given type [#122](https://github.com/SharePoint/office365-cli/issues/122)
- [spo list list](../cmd/spo/list/list-list.md) - lists all available list in the specified site [#198](https://github.com/SharePoint/office365-cli/issues/198)
- [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md) - lists hub sites in the current tenant [#91](https://github.com/SharePoint/office365-cli/issues/91)
- [spo hubsite get](../cmd/spo/hubsite/hubsite-get.md) - gets information about the specified hub site [#92](https://github.com/SharePoint/office365-cli/issues/92)
- [spo hubsite register](../cmd/spo/hubsite/hubsite-register.md) - registers the specified site collection as a hub site [#94](https://github.com/SharePoint/office365-cli/issues/94)
- [spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md) - unregisters the specified site collection as a hub site [#95](https://github.com/SharePoint/office365-cli/issues/95)
- [spo hubsite set](../cmd/spo/hubsite/hubsite-set.md) - updates properties of the specified hub site [#96](https://github.com/SharePoint/office365-cli/issues/96)
- [spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md) - connects the specified site collection to the given hub site [#97](https://github.com/SharePoint/office365-cli/issues/97)
- [spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md) - disconnects the specifies site collection from its hub site [#98](https://github.com/SharePoint/office365-cli/issues/98)
- [spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md) - grants permissions to join the hub site for one or more principals [#99](https://github.com/SharePoint/office365-cli/issues/99)
- [spo hubsite rights revoke](../cmd/spo/hubsite/hubsite-rights-revoke.md) - revokes rights to join sites to the specified hub site for one or more principals [#100](https://github.com/SharePoint/office365-cli/issues/100)
- [spo customaction set](../cmd/spo/customaction/customaction-set.md) - updates a user custom action for site or site collection [#212](https://github.com/SharePoint/office365-cli/issues/212)

### Changes

- fixed issue with prompts in non-interactive mode [#142](https://github.com/SharePoint/office365-cli/issues/142)
- added information about the current user to status commands [#202](https://github.com/SharePoint/office365-cli/issues/202)
- fixed issue with completing input that doesn't match commands [#222](https://github.com/SharePoint/office365-cli/issues/222)
- fixed issue with escaping numeric input [#226](https://github.com/SharePoint/office365-cli/issues/226)
- changed the [aad oauth2grant list](../cmd/aad/oauth2grant/oauth2grant-list.md), [spo app list](../cmd/spo/app/app-list.md), [spo customaction list](../cmd/spo/customaction/customaction-list.md), [spo site list](../cmd/spo/site/site-list.md) commands to list all properties for output type JSON [#232](https://github.com/SharePoint/office365-cli/issues/232), [#233](https://github.com/SharePoint/office365-cli/issues/233), [#234](https://github.com/SharePoint/office365-cli/issues/234), [#235](https://github.com/SharePoint/office365-cli/issues/235)
- fixed issue with generating clink completion file [#252](https://github.com/SharePoint/office365-cli/issues/252)
- added [user guide](../user-guide/installing-cli.md) [#236](https://github.com/SharePoint/office365-cli/issues/236), [#237](https://github.com/SharePoint/office365-cli/issues/237), [#238](https://github.com/SharePoint/office365-cli/issues/238), [#239](https://github.com/SharePoint/office365-cli/issues/239)

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