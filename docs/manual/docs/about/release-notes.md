# Release notes

## [v1.18.0](https://github.com/pnp/office365-cli/releases/tag/v1.17.0)

### New commands

**Microsoft Graph:**

- [graph o365group renew](../cmd/graph/o365group/o365group-renew.md) - renews Office 365 group's expiration [#870](https://github.com/pnp/office365-cli/issues/870)
- [graph o365group user remove](../cmd/graph/o365group/o365group-user-remove.md) - removes the specified user from specified Office 365 Group or Microsoft Teams team [#846](https://github.com/pnp/office365-cli/issues/846)

### Changes

- centralized executing HTTP requests solving [#888](https://github.com/pnp/office365-cli/issues/888)

## [v1.17.0](https://github.com/pnp/office365-cli/releases/tag/v1.17.0)

### New commands

**SharePoint Online:**

- [spo contenttype field remove](../cmd/spo/contenttype/contenttype-field-remove.md) - removes a column from a site- or list content type [#673](https://github.com/pnp/office365-cli/issues/673)
- [spo mail send](../cmd/spo/mail/mail-send.md) - sends an e-mail from SharePoint [#753](https://github.com/pnp/office365-cli/issues/753)

**Microsoft Graph:**

- [graph teams archive](../cmd/graph/teams/teams-archive.md) - archives specified Microsoft Teams team [#899](https://github.com/pnp/office365-cli/issues/899)
- [graph teams channel get](../cmd/graph/teams/teams-channel-get.md) - gets information about the specific Microsoft Teams team channel [#808](https://github.com/pnp/office365-cli/issues/808)
- [graph teams messagingsettings set](../cmd/graph/teams/teams-messagingsettings-set.md) - updates messaging settings of a Microsoft Teams team [#820](https://github.com/pnp/office365-cli/issues/820)
- [graph teams remove](../cmd/graph/teams/teams-remove.md) - removes the specified Microsoft Teams team [#813](https://github.com/pnp/office365-cli/issues/813)
- [graph teams unarchive](../cmd/graph/teams/teams-unarchive.md) - restores an archived Microsoft Teams team [#900](https://github.com/pnp/office365-cli/issues/900)

### Changes

- updated documentation on using custom AAD app [#895](https://github.com/pnp/office365-cli/issues/895)
- added validation for Teams channel IDs [#909](https://github.com/pnp/office365-cli/issues/909)
- fixed the 'spo page clientsidewebpart add' command [#913](https://github.com/pnp/office365-cli/issues/913)
- fixed typo in the 'spo tenant settings set' command options [#923](https://github.com/pnp/office365-cli/issues/923)
- updated commands to use MS Graph v1.0 endpoint [#865](https://github.com/pnp/office365-cli/issues/865)
- added support for upgrading projects built using SharePoint Framework v1.8.0 [#932](https://github.com/pnp/office365-cli/issues/932)

## [v1.16.0](https://github.com/pnp/office365-cli/releases/tag/v1.16.0)

### New commands

**SharePoint Online:**

- [spo listitem record declare](../cmd/spo/listitem/listitem-record-declare.md) - declares the specified list item as a record [#769](https://github.com/pnp/office365-cli/issues/769)

**Microsoft Graph:**

- [graph o365group user add](../cmd/graph/o365group/o365group-user-add.md) - adds user to specified Office 365 Group or Microsoft Teams team [#847](https://github.com/pnp/office365-cli/issues/847)
- [graph schemaextension add](../cmd/graph/schemaextension/schemaextension-add.md) - creates a Microsoft Graph schema extension [#13](https://github.com/pnp/office365-cli/issues/13)
- [graph teams add](../cmd/graph/teams/teams-add.md) - adds a new Microsoft Teams team [#615](https://github.com/pnp/office365-cli/issues/615)
- [graph teams app uninstall](../cmd/graph/teams/teams-app-uninstall.md) - uninstalls an app from a Microsoft Team team [#843](https://github.com/pnp/office365-cli/issues/843)
- [graph teams channel set](../cmd/graph/teams/teams-channel-set.md) - updates properties of the specified channel in the given Microsoft Teams team [#816](https://github.com/pnp/office365-cli/issues/816)
- [graph teams guestsettings set](../cmd/graph/teams/teams-guestsettings-set.md) - updates guest settings of a Microsoft Teams team [#818](https://github.com/pnp/office365-cli/issues/818)
- [graph teams tab list](../cmd/graph/teams/teams-tab-list.md) - lists tabs in the specified Microsoft Teams channel [#849](https://github.com/pnp/office365-cli/issues/849)

### Changes

- extended 'graph teams app list' [#859](https://github.com/pnp/office365-cli/issues/859)
- added 'spo site groupify' alias [#873](https://github.com/pnp/office365-cli/issues/873)
- fixed the 'spo page section add' command [#908](https://github.com/pnp/office365-cli/issues/908)
- fixed the 'spo page header set' command [#911](https://github.com/pnp/office365-cli/issues/911)

## [v1.15.0](https://github.com/pnp/office365-cli/releases/tag/v1.15.0)

### New commands

**SharePoint Online:**

- [spo field remove](../cmd/spo/field/field-remove.md) - removes the specified list- or site column [#738](https://github.com/pnp/office365-cli/issues/738)
- [spo listitem record undeclare](../cmd/spo/listitem/listitem-record-undeclare.md) - undeclares list item as a record [#770](https://github.com/pnp/office365-cli/issues/770)
- [spo web reindex](../cmd/spo/web/web-reindex.md) - requests reindexing the specified subsite [#822](https://github.com/pnp/office365-cli/issues/822)

**Microsoft Graph:**

- [graph teams app install](../cmd/graph/teams/teams-app-install.md) - installs an app from the catalog to a Microsoft Teams team [#842](https://github.com/pnp/office365-cli/issues/842)
- [graph teams funsettings list](../cmd/graph/teams/teams-funsettings-list.md) - lists fun settings for the specified Microsoft Teams team [#809](https://github.com/pnp/office365-cli/issues/809)
- [graph teams guestsettings list](../cmd/graph/teams/teams-guestsettings-list.md) - lists guests settings for a Microsoft Teams team [#810](https://github.com/pnp/office365-cli/issues/810)
- [graph teams membersettings list](../cmd/graph/teams/teams-membersettings-list.md) - lists member settings for a Microsoft Teams team [#811](https://github.com/pnp/office365-cli/issues/811)
- [graph teams membersettings set](../cmd/graph/teams/teams-membersettings-set.md) - updates member settings of a Microsoft Teams team [#819](https://github.com/pnp/office365-cli/issues/819)
- [graph teams messagingsettings list](../cmd/graph/teams/teams-messagingsettings-list.md) - lists messaging settings for a Microsoft Teams team [#812](https://github.com/pnp/office365-cli/issues/812)

### Changes

- fixed ID of the FN002009 SPFx project upgrade rule [#854](https://github.com/pnp/office365-cli/issues/854)
- fixed issue with updating the header of non-en-US pages [#851](https://github.com/pnp/office365-cli/issues/851)
- added support for upgrading projects built using SharePoint Framework v1.7.1 [#848](https://github.com/pnp/office365-cli/issues/848)

## [v1.14.0](https://github.com/pnp/office365-cli/releases/tag/v1.14.0)

### New commands

**SharePoint Online:**

- [spo list label get](../cmd/spo/list/list-label-get.md) - gets label set on the specified list [#773](https://github.com/pnp/office365-cli/issues/773)
- [spo list label set](../cmd/spo/list/list-label-set.md) - sets classification label on the specified list [#772](https://github.com/pnp/office365-cli/issues/772)
- [spo list view field add](../cmd/spo/list/list-view-field-add.md) - adds the specified field to list view [#735](https://github.com/pnp/office365-cli/issues/735)
- [spo list view field remove](../cmd/spo/list/list-view-field-remove.md) - removes the specified field from list view [#736](https://github.com/pnp/office365-cli/issues/736)
- [spo site inplacerecordsmanagement set](../cmd/spo/site/site-inplacerecordsmanagement-set.md) - activates or deactivates in-place records management for a site collection [#774](https://github.com/pnp/office365-cli/issues/774)
- [spo sitedesign run list](../cmd/spo/sitedesign/sitedesign-run-list.md) - lists information about site designs applied to the specified site [#779](https://github.com/pnp/office365-cli/issues/779)
- [spo sitedesign run status get](../cmd/spo/sitedesign/sitedesign-run-status-get.md) - gets information about the site scripts executed for the specified site design [#780](https://github.com/pnp/office365-cli/issues/780)
- [spo sitedesign task get](../cmd/spo/sitedesign/sitedesign-task-get.md) - gets information about the specified site design scheduled for execution [#782](https://github.com/pnp/office365-cli/issues/782)
- [spo sitedesign task list](../cmd/spo/sitedesign/sitedesign-task-list.md) - lists site designs scheduled for execution on the specified site [#781](https://github.com/pnp/office365-cli/issues/781)

**Microsoft Graph:**

- [graph teams app list](../cmd/graph/teams/teams-app-list.md) - lists apps from the Microsoft Teams app catalog [#826](https://github.com/pnp/office365-cli/issues/826)
- [graph teams app publish](../cmd/graph/teams/teams-app-publish.md) - publishes Teams app to the organization's app catalog [#824](https://github.com/pnp/office365-cli/issues/824)
- [graph teams app remove](../cmd/graph/teams/teams-app-remove.md) - removes a Teams app from the organization's app catalog [#825](https://github.com/pnp/office365-cli/issues/825)
- [graph teams app update](../cmd/graph/teams/teams-app-update.md) - updates Teams app in the organization's app catalog [#827](https://github.com/pnp/office365-cli/issues/827)
- [graph teams channel list](../cmd/graph/teams/teams-channel-list.md) - lists channels in the specified Microsoft Teams team [#586](https://github.com/pnp/office365-cli/issues/586)
- [graph teams user remove](../cmd/graph/o365group/o365group-user-remove.md) - removes the specified user from the specified Microsoft Teams team [#757](https://github.com/pnp/office365-cli/issues/757)
- [graph teams user set](../cmd/graph/teams/teams-user-set.md) - updates role of the specified user in the given Microsoft Teams team [#760](https://github.com/pnp/office365-cli/issues/760)

### Changes

- updated 'spo list webhook list' parameters [#747](https://github.com/pnp/office365-cli/issues/747)
- updated 'azmgmt flow list' to support paged content [#776](https://github.com/pnp/office365-cli/issues/776)
- added fieldTitle, listId and listUrl options to 'spo file get' [#754](https://github.com/pnp/office365-cli/issues/754)
- extended 'spo sitedesign apply' with large site designs [#714](https://github.com/pnp/office365-cli/issues/714)
- added support for dynamic data [#751](https://github.com/pnp/office365-cli/issues/751)
- extended 'spo web set' with modern UI options [#798](https://github.com/pnp/office365-cli/issues/798)

## [v1.13.0](https://github.com/pnp/office365-cli/releases/tag/v1.13.0)

### New commands

**SharePoint Online:**

- [spo feature list](../cmd/spo/feature/feature-list.md) - lists Features activated in the specified site or site collection [#677](https://github.com/pnp/office365-cli/issues/677)
- [spo file move](../cmd/spo/file/file-move.md) - moves a file to another location [#671](https://github.com/pnp/office365-cli/issues/671)
- [spo list view list](../cmd/spo/list/list-view-list.md) - lists views configured on the specified list [#732](https://github.com/pnp/office365-cli/issues/732)
- [spo list sitescript get](../cmd/spo/list/list-sitescript-get.md) - extracts a site script from a SharePoint list [#713](https://github.com/pnp/office365-cli/issues/713)
- [spo list view get](../cmd/spo/list/list-view-get.md) - gets information about specific list view [#730](https://github.com/pnp/office365-cli/issues/730)
- [spo list view remove](../cmd/spo/list/list-view-remove.md) - deletes the specified view from the list [#731](https://github.com/pnp/office365-cli/issues/731)

**Microsoft Graph:**

- [graph teams message list](../cmd/graph/teams/teams-message-list.md) - lists all messages from a channel in a Microsoft Teams team [#588](https://github.com/pnp/office365-cli/issues/588)
- [graph teams user add](../cmd/graph/o365group/o365group-user-add.md) - adds user to the specified Microsoft Teams team [#690](https://github.com/pnp/office365-cli/issues/690)
- [graph teams user list](../cmd/graph/teams/teams-user-list.md) - lists users for the specified Microsoft Teams team [#689](https://github.com/pnp/office365-cli/issues/689)

### Changes

- added support for specifying language when creating site [#728](https://github.com/pnp/office365-cli/issues/728)
- fixed bug in setting client-side web part order [#712](https://github.com/pnp/office365-cli/issues/712)
- added support for authentication using certificate [#389](https://github.com/pnp/office365-cli/issues/389)
- renamed 'graph teams channel message get' to 'graph teams message get'
- extended 'spo folder copy' with support for schema mismatch [#706](https://github.com/pnp/office365-cli/pull/706)
- extended 'spo file copy' with support for schema mismatch [#705](https://github.com/pnp/office365-cli/pull/705)
- updated showing scope in 'spo customaction list' [#742](https://github.com/pnp/office365-cli/issues/742)
- extended 'spo hubsite list' with info about associated sites [#709](https://github.com/pnp/office365-cli/pull/709)
- added support for SPO-D URLs solving [#759](https://github.com/pnp/office365-cli/pull/759)

## [v1.12.0](https://github.com/pnp/office365-cli/releases/tag/v1.12.0)

### New commands

**SharePoint Online:**

- [spo folder move](../cmd/spo/folder/folder-move.md) - moves a folder to another location [#672](https://github.com/pnp/office365-cli/issues/672)
- [spo page text add](../cmd/spo/page/page-text-add.md) - adds text to a modern page [#365](https://github.com/pnp/office365-cli/issues/365)

### Changes

- added support for site collection app catalog in the spo app install, -retract, -uninstall and -upgrade commands [#405](https://github.com/pnp/office365-cli/issues/405)
- fixed bug with caching tokens for SPO commands [#719](https://github.com/pnp/office365-cli/issues/719)

## [v1.11.0](https://github.com/pnp/office365-cli/releases/tag/v1.11.0)

### New commands

**SharePoint Online:**

- [spo list webhook add](../cmd/spo/list/list-webhook-add.md) - adds a new webhook to the specified list [#652](https://github.com/pnp/office365-cli/issues/652)
- [spo page header set](../cmd/spo/page/page-header-set.md) - sets modern page header [#697](https://github.com/pnp/office365-cli/issues/697)

### Changes

- added support for setting page title [#693](https://github.com/pnp/office365-cli/issues/693)
- added support for adding child navigation nodes [#695](https://github.com/pnp/office365-cli/issues/695)
- added support for specifying web part data and fixed web parts lookup [#701](https://github.com/pnp/office365-cli/issues/701), [#703](https://github.com/pnp/office365-cli/issues/703)
- removed treating values of unknown options as numbers [#702](https://github.com/pnp/office365-cli/issues/702)
- added support for site collection app catalog in the spo app add, -deploy, -get, -list and -remove commands [#405](https://github.com/pnp/office365-cli/issues/405) (partially)
- added support for upgrading projects built using SharePoint Framework v1.7.0 [#716](https://github.com/pnp/office365-cli/pull/716)

## [v1.10.0](https://github.com/pnp/office365-cli/releases/tag/v1.10.0)

### New commands

**SharePoint Online:**

- [spo field set](../cmd/spo/field/field-set.md) - updates existing list or site column [#661](https://github.com/pnp/office365-cli/issues/661)
- [spo file add](../cmd/spo/file/file-add.md) - uploads file to the specified folder [#283](https://github.com/pnp/office365-cli/issues/283)
- [spo list contenttype add](../cmd/spo/list/list-contenttype-add.md) - adds content type to list [#594](https://github.com/pnp/office365-cli/issues/594)
- [spo list contenttype list](../cmd/spo/list/list-contenttype-list.md) - lists content types configured on the list [#595](https://github.com/pnp/office365-cli/issues/595)
- [spo list contenttype remove](../cmd/spo/list/list-contenttype-remove.md) - removes content type from list [#668](https://github.com/pnp/office365-cli/issues/668)
- [spo list view set](../cmd/spo/list/list-view-set.md) - updates existing list view [#662](https://github.com/pnp/office365-cli/issues/662)
- [spo list webhook remove](../cmd/spo/list/list-webhook-remove.md) - removes the specified webhook from the list [#650](https://github.com/pnp/office365-cli/issues/650)
- [spo list webhook set](../cmd/spo/list/list-webhook-set.md) - updates the specified webhook [#651](https://github.com/pnp/office365-cli/issues/651)
- [spo search](../cmd/spo/search/search.md) - executes a search query [#316](https://github.com/pnp/office365-cli/issues/316)
- [spo serviceprincipal grant add](../cmd/spo/serviceprincipal/serviceprincipal-grant-add.md) - grants the service principal permission to the specified API [#590](https://github.com/pnp/office365-cli/issues/590)

**Microsoft Graph:**

- [graph siteclassification set](../cmd/graph/siteclassification/siteclassification-set.md) - updates site classification configuration [#304](https://github.com/pnp/office365-cli/issues/304)
- [graph teams channel message get](../cmd/graph/teams/teams-message-get.md) - retrieves a message from a channel in a Microsoft Teams team [#589](https://github.com/pnp/office365-cli/issues/589)

### Changes

- added support for adding child terms [#686](https://github.com/pnp/office365-cli/issues/686)

## [v1.9.0](https://github.com/pnp/office365-cli/releases/tag/v1.9.0)

### Changes

- added support for upgrading projects built using SharePoint Framework v1.6.0 [#663](https://github.com/pnp/office365-cli/issues/663)

## [v1.8.0](https://github.com/pnp/office365-cli/releases/tag/v1.8.0)

### New commands

**SharePoint Online:**

- [spo list webhook get](../cmd/spo/list/list-webhook-get.md) - gets information about the specific webhook [#590](https://github.com/pnp/office365-cli/issues/590)
- [spo tenant settings set](../cmd/spo/tenant/tenant-settings-set.md) - sets tenant global settings [#549](https://github.com/pnp/office365-cli/issues/549)
- [spo term add](../cmd/spo/term/term-add.md) - adds taxonomy term [#605](https://github.com/pnp/office365-cli/issues/605)
- [spo term get](../cmd/spo/term/term-get.md) - gets information about the specified taxonomy term [#604](https://github.com/pnp/office365-cli/issues/604)
- [spo term list](../cmd/spo/term/term-list.md) - lists taxonomy terms from the given term set [#603](https://github.com/pnp/office365-cli/issues/603)
- [spo term group add](../cmd/spo/term/term-group-add.md) - adds taxonomy term group [#598](https://github.com/pnp/office365-cli/issues/598)
- [spo term set add](../cmd/spo/term/term-set-add.md) - adds taxonomy term set [#602](https://github.com/pnp/office365-cli/issues/602)
- [spo term set get](../cmd/spo/term/term-set-get.md) - gets information about the specified taxonomy term set [#601](https://github.com/pnp/office365-cli/issues/601)
- [spo term set list](../cmd/spo/term/term-set-list.md) - lists taxonomy term sets from the given term group [#600](https://github.com/pnp/office365-cli/issues/600)

**Microsoft Graph:**

- [graph siteclassification disable](../cmd/graph/siteclassification/siteclassification-disable.md) - disables site classification [#302](https://github.com/pnp/office365-cli/issues/302)
- [graph siteclassification enable](../cmd/graph/siteclassification/siteclassification-enable.md) - enables site classification [#301](https://github.com/pnp/office365-cli/issues/301)
- [graph teams channel add](../cmd/graph/teams/teams-channel-add.md) - adds a channel to the specified Microsoft Teams team [#587](https://github.com/pnp/office365-cli/issues/587)

### Changes

- improved SPFx project upgrade text report [#591](https://github.com/pnp/office365-cli/issues/591)
- updated the 'spo tenant settings list' command [#623](https://github.com/pnp/office365-cli/issues/623)
- changed commands to be lazy-loaded [#624](https://github.com/pnp/office365-cli/issues/624)
- added error codes to the 'spfx project upgrade' command [#630](https://github.com/pnp/office365-cli/issues/630)
- changed vorpal dependency to https [#637](https://github.com/pnp/office365-cli/issues/637)
- added retrieving GuestUsageGuidelinesUrl [#640](https://github.com/pnp/office365-cli/issues/640)

## [v1.7.0](https://github.com/pnp/office365-cli/releases/tag/v1.7.0)

### New commands

**SharePoint Online:**

- [spo list webhook list](../cmd/spo/list/list-webhook-list.md) - lists all webhooks for the specified list [#579](https://github.com/pnp/office365-cli/issues/579)
- [spo listitem list](../cmd/spo/listitem/listitem-list.md) - gets a list of items from the specified list [#268](https://github.com/pnp/office365-cli/issues/268)
- [spo page column get](../cmd/spo/page/page-column-get.md) - get information about a specific column of a modern page [#412](https://github.com/pnp/office365-cli/issues/412)
- [spo page remove](../cmd/spo/page/page-remove.md) - removes a modern page [#363](https://github.com/pnp/office365-cli/issues/363)
- [spo page section add](../cmd/spo/page/page-section-add.md) - adds section to modern page [#364](https://github.com/pnp/office365-cli/issues/364)
- [spo site classic remove](../cmd/spo/site/site-classic-remove.md) - removes the specified site [#125](https://github.com/pnp/office365-cli/issues/125)
- [spo tenant settings list](../cmd/spo/tenant/tenant-settings-list.md) - lists the global tenant settings [#548](https://github.com/pnp/office365-cli/issues/548)
- [spo term group get](../cmd/spo/term/term-group-get.md) - gets information about the specified taxonomy term group [#597](https://github.com/pnp/office365-cli/issues/597)
- [spo term group list](../cmd/spo/term/term-group-list.md) - lists taxonomy term groups [#596](https://github.com/pnp/office365-cli/issues/596)

**Microsoft Graph:**

- [graph groupsetting remove](../cmd/graph/groupsetting/groupsetting-remove.md) - removes the particular group setting [#452](https://github.com/pnp/office365-cli/pull/452)
- [graph groupsetting set](../cmd/graph/groupsetting/groupsetting-set.md) - removes the particular group setting [#451](https://github.com/pnp/office365-cli/pull/451)

**Azure Management Service:**

- [azmgmt flow export](../cmd/azmgmt/flow/flow-export.md) - exports the specified Microsoft Flow as a file [#383](https://github.com/pnp/office365-cli/issues/383)
- [azmgmt flow run get](../cmd/azmgmt/flow/flow-run-get.md) - gets information about a specific run of the specified Microsoft Flow [#400](https://github.com/pnp/office365-cli/issues/400)
- [azmgmt flow run list](../cmd/azmgmt/flow/flow-run-list.md) - lists runs of the specified Microsoft Flow [#399](https://github.com/pnp/office365-cli/issues/399)

### Changes

- added support for upgrading projects built using SharePoint Framework v1.5.1 [#569](https://github.com/pnp/office365-cli/issues/569)
- added support for setting debug and verbose mode using an environment variable [#54](https://github.com/pnp/office365-cli/issues/54)
- extended the 'spo cdn set' command, solving. Added support for managing both CDNs. Added support for enabling CDNs without provisioning default origins [#230](https://github.com/pnp/office365-cli/issues/230)
- fixed bug in reporting SPFx project upgrade findings solving [#582](https://github.com/pnp/office365-cli/issues/582)
- fixed upgrade SPFx 1.6.0 FN012012 always returns a finding [#580](https://github.com/pnp/office365-cli/issues/580)
- combined npm commands in SPFx project upgrade summary solving [#508](https://github.com/pnp/office365-cli/issues/508)
- renamed 'connect' commands to 'login' [#574](https://github.com/pnp/office365-cli/issues/574)
- updated docs on escaping objectId in aad oauth2grant set and remove [#606](https://github.com/pnp/office365-cli/issues/606)
- added 'npm dedupe' SPFx project upgrade rule [#612](https://github.com/pnp/office365-cli/issues/612)

## [v1.6.0](https://github.com/pnp/office365-cli/releases/tag/v1.6.0)

### New commands

**SharePoint Online:**

- [spo contenttype field set](../cmd/spo/contenttype/contenttype-field-set.md) - adds or updates a site column reference in a site content type [#520](https://github.com/pnp/office365-cli/issues/520)
- [spo page section get](../cmd/spo/page/page-section-get.md) - gets information about the specified modern page section [#410](https://github.com/pnp/office365-cli/issues/410)
- [spo page section list](../cmd/spo/page/page-section-list.md) - lists sections in the specific modern page [#409](https://github.com/pnp/office365-cli/issues/409)

**Microsoft Graph:**

- [graph teams list](../cmd/graph/teams/teams-list.md) - lists Microsoft Teams in the current tenant [#558](https://github.com/pnp/office365-cli/pull/558)

### Changes

- added support for upgrading projects built using SharePoint Framework v1.1.3 [#485](https://github.com/pnp/office365-cli/issues/485)
- added support for upgrading projects built using SharePoint Framework v1.1.1 [#487](https://github.com/pnp/office365-cli/issues/487)
- added support for upgrading projects built using SharePoint Framework v1.1.0 [#488](https://github.com/pnp/office365-cli/issues/488)
- added support for upgrading projects built using SharePoint Framework v1.0.2 [#537](https://github.com/pnp/office365-cli/issues/537)
- added support for upgrading projects built using SharePoint Framework v1.0.1 [#536](https://github.com/pnp/office365-cli/issues/536)
- added support for upgrading projects built using SharePoint Framework v1.0.0 [#535](https://github.com/pnp/office365-cli/issues/535)
- fixed created content type have different ID than specified [#550](https://github.com/pnp/office365-cli/issues/550)

## [v1.5.0](https://github.com/pnp/office365-cli/releases/tag/v1.5.0)

### New commands

**SharePoint Online:**

- [spo contenttype add](../cmd/spo/contenttype/contenttype-add.md) - adds a new list or site content type [#519](https://github.com/pnp/office365-cli/issues/519)
- [spo contenttype get](../cmd/spo/contenttype/contenttype-get.md) - retrieves information about the specified list or site content type [#532](https://github.com/pnp/office365-cli/issues/532)
- [spo field add](../cmd/spo/field/field-add.md) - adds a new list or site column using the CAML field definition [#518](https://github.com/pnp/office365-cli/issues/518)
- [spo field get](../cmd/spo/field/field-get.md) - retrieves information about the specified list or site column [#528](https://github.com/pnp/office365-cli/issues/528)
- [spo navigation node add](../cmd/spo/navigation/navigation-node-add.md) - adds a navigation node to the specified site navigation [#521](https://github.com/pnp/office365-cli/issues/521)
- [spo navigation node list](../cmd/spo/navigation/navigation-node-list.md) - lists nodes from the specified site navigation [#522](https://github.com/pnp/office365-cli/issues/522)
- [spo navigation node remove](../cmd/spo/navigation/navigation-node-remove.md) - removes the specified navigation node [#523](https://github.com/pnp/office365-cli/issues/523)
- [spo page clientsidewebpart add](../cmd/spo/page/page-clientsidewebpart-add.md) - adds a client-side web part to a modern page [#366](https://github.com/pnp/office365-cli/issues/366)
- [spo page column list](../cmd/spo/page/page-column-list.md) - lists columns in the specific section of a modern page [#411](https://github.com/pnp/office365-cli/issues/411)
- [spo web set](../cmd/spo/web/web-set.md) - updates subsite properties [#191](https://github.com/pnp/office365-cli/issues/191)

### Changes

- fixed exit code on error in the 'spo site add' command [#511](https://github.com/pnp/office365-cli/issues/511)
- Added support for retrieving apps by their name [#516](https://github.com/pnp/office365-cli/issues/516)

## [v1.4.0](https://github.com/pnp/office365-cli/releases/tag/v1.4.0)

### New commands

**SharePoint Online:**

- [spo file checkin](../cmd/spo/file/file-checkin.md) - checks in specified file [#284](https://github.com/pnp/office365-cli/issues/284)
- [spo file checkout](../cmd/spo/file/file-checkout.md) - checks out specified file [#285](https://github.com/pnp/office365-cli/issues/285)
- [spo folder rename](../cmd/spo/folder/folder-rename.md) - renames a folder [#429](https://github.com/pnp/office365-cli/issues/429)
- [spo listitem get](../cmd/spo/listitem/listitem-get.md) - gets a list item from the specified list [#269](https://github.com/pnp/office365-cli/issues/269)
- [spo listitem set](../cmd/spo/listitem/listitem-set.md) - updates a list item in the specified list [#271](https://github.com/pnp/office365-cli/issues/271)

**SharePoint Framework:**

- [spfx project upgrade](../cmd/spfx/project/project-upgrade.md) - upgrades SharePoint Framework project to the specified version [#471](https://github.com/pnp/office365-cli/issues/471)

### Changes

- refactored to return non-zero error code on error [#468](https://github.com/pnp/office365-cli/issues/468)
- fixed adding item to list referenced by id [#473](https://github.com/pnp/office365-cli/issues/473)
- added support for upgrading projects built using SharePoint Framework v1.4.0 [#478](https://github.com/pnp/office365-cli/issues/478)
- added support for upgrading projects built using SharePoint Framework v1.3.4 [#479](https://github.com/pnp/office365-cli/issues/479)
- added support for upgrading projects built using SharePoint Framework v1.3.2 [#481](https://github.com/pnp/office365-cli/issues/481)
- added support for upgrading projects built using SharePoint Framework v1.3.1 [#482](https://github.com/pnp/office365-cli/issues/482)
- added support for upgrading projects built using SharePoint Framework v1.3.0 [#483](https://github.com/pnp/office365-cli/issues/483)
- added support for upgrading projects built using SharePoint Framework v1.2.0 [#484](https://github.com/pnp/office365-cli/issues/484)
- clarified usage of the [spo file get](../cmd/spo/file/file-get.md) command [#497](https://github.com/pnp/office365-cli/pull/497)
- added support for upgrading projects built using SharePoint Framework v1.5.0 [#505](https://github.com/pnp/office365-cli/issues/505)

## [v1.3.0](https://github.com/pnp/office365-cli/releases/tag/v1.3.0)

### New commands

**SharePoint Online:**

- [spo file copy](../cmd/spo/file/file-copy.md) - copies a file to another location [#286](https://github.com/pnp/office365-cli/issues/286)
- [spo folder add](../cmd/spo/folder/folder-add.md) - creates a folder within a parent folder [#425](https://github.com/pnp/office365-cli/issues/425)
- [spo folder copy](../cmd/spo/folder/folder-copy.md) - copies a folder to another location [#424](https://github.com/pnp/office365-cli/issues/424)
- [spo folder get](../cmd/spo/folder/folder-get.md) - gets information about the specified folder [#427](https://github.com/pnp/office365-cli/issues/427)
- [spo folder list](../cmd/spo/folder/folder-list.md) - returns all folders under the specified parent folder [#428](https://github.com/pnp/office365-cli/issues/428)
- [spo folder remove](../cmd/spo/folder/folder-remove.md) - deletes the specified folder [#426](https://github.com/pnp/office365-cli/issues/426)
- [spo hidedefaultthemes get](../cmd/spo/hidedefaultthemes/hidedefaultthemes-get.md) - gets the current value of the HideDefaultThemes setting [#341](https://github.com/pnp/office365-cli/issues/341)
- [spo hidedefaultthemes set](../cmd/spo/hidedefaultthemes/hidedefaultthemes-set.md) - sets the value of the HideDefaultThemes setting [#342](https://github.com/pnp/office365-cli/issues/342)
- [spo site o365group set](../cmd/spo/site/site-o365group-set.md) - connects site collection to an Office 365 Group [#431](https://github.com/pnp/office365-cli/issues/431)
- [spo theme apply](../cmd/spo/theme/theme-apply.md) - applies theme to the specified site [#343](https://github.com/pnp/office365-cli/issues/343)

**Microsoft Graph:**

- [graph groupsetting add](../cmd/graph/groupsetting/groupsetting-add.md) - creates a group setting [#443](https://github.com/pnp/office365-cli/issues/443)
- [graph groupsetting get](../cmd/graph/groupsetting/groupsetting-get.md) - gets information about the particular group setting [#450](https://github.com/pnp/office365-cli/issues/450)
- [graph groupsetting list](../cmd/graph/groupsetting/groupsetting-list.md) - lists Azure AD group settings [#449](https://github.com/pnp/office365-cli/issues/449)
- [graph groupsettingtemplate get](../cmd/graph/groupsettingtemplate/groupsettingtemplate-get.md) - gets information about the specified Azure AD group settings template [#442](https://github.com/pnp/office365-cli/issues/442)
- [graph groupsettingtemplate list](../cmd/graph/groupsettingtemplate/groupsettingtemplate-list.md) - lists Azure AD group settings templates [#441](https://github.com/pnp/office365-cli/issues/441)
- [graph user sendmail](../cmd/graph/user/user-sendmail.md) - sends e-mail on behalf of the current user [#328](https://github.com/pnp/office365-cli/issues/328)

### Changes

- added support for re-consenting the AAD app [#421](https://github.com/pnp/office365-cli/issues/421)
- added update notification [#200](https://github.com/pnp/office365-cli/issues/200)
- extended the 'spo app deploy' command to support specifying app using its name [#404](https://github.com/pnp/office365-cli/issues/404)
- extended the 'spo app add' command to return the information about the added app [#463](https://github.com/pnp/office365-cli/issues/463)

## [v1.2.0](https://github.com/pnp/office365-cli/releases/tag/v1.2.0)

### New commands

**SharePoint Online:**

- [spo file remove](../cmd/spo/file/file-remove.md) - removes the specified file [#287](https://github.com/pnp/office365-cli/issues/287)
- [spo hubsite data get](../cmd/spo/hubsite/hubsite-data-get.md) - gets hub site data for the specified site [#394](https://github.com/pnp/office365-cli/issues/394)
- [spo hubsite theme sync](../cmd/spo/hubsite/hubsite-theme-sync.md) - applies any theme updates from the parent hub site [#401](https://github.com/pnp/office365-cli/issues/401)
- [spo listitem add](../cmd/spo/listitem/listitem-add.md) - creates a list item in the specified list [#270](https://github.com/pnp/office365-cli/issues/270)
- [spo listitem remove](../cmd/spo/listitem/listitem-remove.md) - removes the specified list item [#272](https://github.com/pnp/office365-cli/issues/272)
- [spo page control get](../cmd/spo/page/page-control-get.md) - gets information about the specific control on a modern page [#414](https://github.com/pnp/office365-cli/issues/414)
- [spo page control list](../cmd/spo/page/page-control-list.md) - lists controls on the specific modern page [#413](https://github.com/pnp/office365-cli/issues/413)
- [spo page get](../cmd/spo/page/page-get.md) - gets information about the specific modern page [#360](https://github.com/pnp/office365-cli/issues/360)
- [spo propertybag set](../cmd/spo/propertybag/propertybag-set.md) - sets the value of the specified property in the property bag [#393](https://github.com/pnp/office365-cli/issues/393)
- [spo web clientsidewebpart list](../cmd/spo/web/web-clientsidewebpart-list.md) - lists available client-side web parts [#367](https://github.com/pnp/office365-cli/issues/367)

**Microsoft Graph:**

- [graph user get](../cmd/graph/user/user-get.md) - gets information about the specified user [#326](https://github.com/pnp/office365-cli/issues/326)
- [graph user list](../cmd/graph/user/user-list.md) - lists users matching specified criteria [#327](https://github.com/pnp/office365-cli/issues/327)

### Changes

- added support for authenticating using credentials solving [#388](https://github.com/pnp/office365-cli/issues/388)

## [v1.1.0](https://github.com/pnp/office365-cli/releases/tag/v1.1.0)

### New commands

**SharePoint Online:**

- [spo file get](../cmd/spo/file/file-get.md) - gets information about the specified file [#282](https://github.com/pnp/office365-cli/issues/282)
- [spo page add](../cmd/spo/page/page-add.md) - creates modern page [#361](https://github.com/pnp/office365-cli/issues/361)
- [spo page list](../cmd/spo/page/page-list.md) - lists all modern pages in the given site [#359](https://github.com/pnp/office365-cli/issues/359)
- [spo page set](../cmd/spo/page/page-set.md) - updates modern page properties [#362](https://github.com/pnp/office365-cli/issues/362)
- [spo propertybag remove](../cmd/spo/propertybag/propertybag-remove.md) - removes specified property from the property bag [#291](https://github.com/pnp/office365-cli/issues/291)
- [spo sitedesign apply](../cmd/spo/sitedesign/sitedesign-apply.md) - applies a site design to an existing site collection [#339](https://github.com/pnp/office365-cli/issues/339)
- [spo theme get](../cmd/spo/theme/theme-get.md) - gets custom theme information [#349](https://github.com/pnp/office365-cli/issues/349)
- [spo theme list](../cmd/spo/theme/theme-list.md) - retrieves the list of custom themes [#332](https://github.com/pnp/office365-cli/issues/332)
- [spo theme remove](../cmd/spo/theme/theme-remove.md) - removes existing theme [#331](https://github.com/pnp/office365-cli/issues/331)
- [spo theme set](../cmd/spo/theme/theme-set.md) - add or update a theme [#330](https://github.com/pnp/office365-cli/issues/330), [#340](https://github.com/pnp/office365-cli/issues/340)
- [spo web get](../cmd/spo/web/web-get.md) - retrieve information about the specified site [#188](https://github.com/pnp/office365-cli/issues/188)

**Microsoft Graph:**

- [graph o365group remove](../cmd/graph/o365group/o365group-remove.md) - removes an Office 365 Group [#309](https://github.com/pnp/office365-cli/issues/309)
- [graph o365group restore](../cmd/graph/o365group/o365group-restore.md) - restores a deleted Office 365 Group [#346](https://github.com/pnp/office365-cli/issues/346)
- [graph siteclassification get](../cmd/graph/siteclassification/siteclassification-get.md) - gets site classification configuration [#303](https://github.com/pnp/office365-cli/issues/303)

**Azure Management Service:**

- [azmgmt login](../cmd/azmgmt/login.md) - log in to the Azure Management Service [#378](https://github.com/pnp/office365-cli/issues/378)
- [azmgmt logout](../cmd/azmgmt/logout.md) - log out from the Azure Management Service [#378](https://github.com/pnp/office365-cli/issues/378)
- [azmgmt status](../cmd/azmgmt/status.md) - shows Azure Management Service login status [#378](https://github.com/pnp/office365-cli/issues/378)
- [azmgmt flow environment get](../cmd/azmgmt/flow/flow-environment-get.md) - gets information about the specified Microsoft Flow environment [#380](https://github.com/pnp/office365-cli/issues/380)
- [azmgmt flow environment list](../cmd/azmgmt/flow/flow-environment-list.md) - lists Microsoft Flow environments in the current tenant [#379](https://github.com/pnp/office365-cli/issues/379)
- [azmgmt flow get](../cmd/azmgmt/flow/flow-get.md) - gets information about the specified Microsoft Flow [#382](https://github.com/pnp/office365-cli/issues/382)
- [azmgmt flow list](../cmd/azmgmt/flow/flow-list.md) - lists Microsoft Flows in the given environment [#381](https://github.com/pnp/office365-cli/issues/381)

### Updated commands

**Microsoft Graph:**

- [graph o365group list](../cmd/graph/o365group/o365group-list.md) - added support for listing deleted Office 365 Groups [#347](https://github.com/pnp/office365-cli/issues/347)

### Changes

- fixed bug in retrieving Office 365 groups in immersive mode solving [#351](https://github.com/pnp/office365-cli/issues/351)

## [v1.0.0](https://github.com/pnp/office365-cli/releases/tag/v1.0.0)

### Breaking changes

- switched to a custom Azure AD application for communicating with Office 365. After installing this version you have to reconnect to Office 365

### New commands

**SharePoint Online:**

- [spo file list](../cmd/spo/file/file-list.md) - lists all available files in the specified folder and site [#281](https://github.com/pnp/office365-cli/issues/281)
- [spo list add](../cmd/spo/list/list-add.md) - creates list in the specified site [#204](https://github.com/pnp/office365-cli/issues/204)
- [spo list remove](../cmd/spo/list/list-remove.md) - removes the specified list [#206](https://github.com/pnp/office365-cli/issues/206)
- [spo list set](../cmd/spo/list/list-set.md) - updates the settings of the specified list [#205](https://github.com/pnp/office365-cli/issues/205)
- [spo customaction clear](../cmd/spo/customaction/customaction-clear.md) - deletes all custom actions in the collection [#231](https://github.com/pnp/office365-cli/issues/231)
- [spo propertybag get](../cmd/spo/propertybag/propertybag-get.md) - gets the value of the specified property from the property bag [#289](https://github.com/pnp/office365-cli/issues/289)
- [spo propertybag list](../cmd/spo/propertybag/propertybag-list.md) - gets property bag values [#288](https://github.com/pnp/office365-cli/issues/288)
- [spo site set](../cmd/spo/site/site-set.md) - updates properties of the specified site [#121](https://github.com/pnp/office365-cli/issues/121)
- [spo site classic add](../cmd/spo/site/site-classic-add.md) - creates new classic site [#123](https://github.com/pnp/office365-cli/issues/123)
- [spo site classic set](../cmd/spo/site/site-classic-set.md) - change classic site settings [#124](https://github.com/pnp/office365-cli/issues/124)
- [spo sitedesign set](../cmd/spo/sitedesign/sitedesign-set.md) - updates a site design with new values [#251](https://github.com/pnp/office365-cli/issues/251)
- [spo tenant appcatalogurl get](../cmd/spo/tenant/tenant-appcatalogurl-get.md) - gets the URL of the tenant app catalog [#315](https://github.com/pnp/office365-cli/issues/315)
- [spo web add](../cmd/spo/web/web-add.md) - create new subsite [#189](https://github.com/pnp/office365-cli/issues/189)
- [spo web list](../cmd/spo/web/web-list.md) - lists subsites of the specified site [#187](https://github.com/pnp/office365-cli/issues/187)
- [spo web remove](../cmd/spo/web/web-remove.md) - delete specified subsite [#192](https://github.com/pnp/office365-cli/issues/192)

**Microsoft Graph:**

- [graph login](../cmd/graph/login.md) - log in to the Microsoft Graph [#10](https://github.com/pnp/office365-cli/issues/10)
- [graph logout](../cmd/graph/logout.md) - log out from the Microsoft Graph [#10](https://github.com/pnp/office365-cli/issues/10)
- [graph status](../cmd/graph/status.md) - shows Microsoft Graph login status [#10](https://github.com/pnp/office365-cli/issues/10)
- [graph o365group add](../cmd/graph/o365group/o365group-add.md) - creates Office 365 Group [#308](https://github.com/pnp/office365-cli/issues/308)
- [graph o365group get](../cmd/graph/o365group/o365group-get.md) - gets information about the specified Office 365 Group [#306](https://github.com/pnp/office365-cli/issues/306)
- [graph o365group list](../cmd/graph/o365group/o365group-list.md) - lists Office 365 Groups in the current tenant [#305](https://github.com/pnp/office365-cli/issues/305)
- [graph o365group set](../cmd/graph/o365group/o365group-set.md) - updates Office 365 Group properties [#307](https://github.com/pnp/office365-cli/issues/307)

### Changes

- fixed bug in logging dates [#317](https://github.com/pnp/office365-cli/issues/317)
- fixed typo in the example of the [spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md) command [#338](https://github.com/pnp/office365-cli/issues/338)

## [v0.5.0](https://github.com/pnp/office365-cli/releases/tag/v0.5.0)

### Breaking changes

- changed the [spo site get](../cmd/spo/site/site-get.md) command to return SPSite properties [#293](https://github.com/pnp/office365-cli/issues/293)

### New commands

**SharePoint Online:**

- [spo sitescript add](../cmd/spo/sitescript/sitescript-add.md) - adds site script for use with site designs [#65](https://github.com/pnp/office365-cli/issues/65)
- [spo sitescript list](../cmd/spo/sitescript/sitescript-list.md) - lists site script available for use with site designs [#66](https://github.com/pnp/office365-cli/issues/66)
- [spo sitescript get](../cmd/spo/sitescript/sitescript-get.md) - gets information about the specified site script [#67](https://github.com/pnp/office365-cli/issues/67)
- [spo sitescript remove](../cmd/spo/sitescript/sitescript-remove.md) - removes the specified site script [#68](https://github.com/pnp/office365-cli/issues/68)
- [spo sitescript set](../cmd/spo/sitescript/sitescript-set.md) - updates existing site script [#216](https://github.com/pnp/office365-cli/issues/216)
- [spo sitedesign add](../cmd/spo/sitedesign/sitedesign-add.md) - adds site design for creating modern sites [#69](https://github.com/pnp/office365-cli/issues/69)
- [spo sitedesign get](../cmd/spo/sitedesign/sitedesign-get.md) - gets information about the specified site design [#86](https://github.com/pnp/office365-cli/issues/86)
- [spo sitedesign list](../cmd/spo/sitedesign/sitedesign-list.md) - lists available site designs for creating modern sites [#85](https://github.com/pnp/office365-cli/issues/85)
- [spo sitedesign remove](../cmd/spo/sitedesign/sitedesign-remove.md) - removes the specified site design [#87](https://github.com/pnp/office365-cli/issues/87)
- [spo sitedesign rights grant](../cmd/spo/sitedesign/sitedesign-rights-grant.md) - grants access to a site design for one or more principals [#88](https://github.com/pnp/office365-cli/issues/88)
- [spo sitedesign rights revoke](../cmd/spo/sitedesign/sitedesign-rights-revoke.md) - revokes access from a site design for one or more principals [#89](https://github.com/pnp/office365-cli/issues/89)
- [spo sitedesign rights list](../cmd/spo/sitedesign/sitedesign-rights-list.md) - gets a list of principals that have access to a site design [#90](https://github.com/pnp/office365-cli/issues/90)
- [spo list get](../cmd/spo/list/list-get.md) - gets information about the specific list [#199](https://github.com/pnp/office365-cli/issues/199)
- [spo customaction remove](../cmd/spo/customaction/customaction-remove.md) - removes the specified custom action [#21](https://github.com/pnp/office365-cli/issues/21)
- [spo site classic list](../cmd/spo/site/site-classic-list.md) - lists sites of the given type [#122](https://github.com/pnp/office365-cli/issues/122)
- [spo list list](../cmd/spo/list/list-list.md) - lists all available list in the specified site [#198](https://github.com/pnp/office365-cli/issues/198)
- [spo hubsite list](../cmd/spo/hubsite/hubsite-list.md) - lists hub sites in the current tenant [#91](https://github.com/pnp/office365-cli/issues/91)
- [spo hubsite get](../cmd/spo/hubsite/hubsite-get.md) - gets information about the specified hub site [#92](https://github.com/pnp/office365-cli/issues/92)
- [spo hubsite register](../cmd/spo/hubsite/hubsite-register.md) - registers the specified site collection as a hub site [#94](https://github.com/pnp/office365-cli/issues/94)
- [spo hubsite unregister](../cmd/spo/hubsite/hubsite-unregister.md) - unregisters the specified site collection as a hub site [#95](https://github.com/pnp/office365-cli/issues/95)
- [spo hubsite set](../cmd/spo/hubsite/hubsite-set.md) - updates properties of the specified hub site [#96](https://github.com/pnp/office365-cli/issues/96)
- [spo hubsite connect](../cmd/spo/hubsite/hubsite-connect.md) - connects the specified site collection to the given hub site [#97](https://github.com/pnp/office365-cli/issues/97)
- [spo hubsite disconnect](../cmd/spo/hubsite/hubsite-disconnect.md) - disconnects the specifies site collection from its hub site [#98](https://github.com/pnp/office365-cli/issues/98)
- [spo hubsite rights grant](../cmd/spo/hubsite/hubsite-rights-grant.md) - grants permissions to join the hub site for one or more principals [#99](https://github.com/pnp/office365-cli/issues/99)
- [spo hubsite rights revoke](../cmd/spo/hubsite/hubsite-rights-revoke.md) - revokes rights to join sites to the specified hub site for one or more principals [#100](https://github.com/pnp/office365-cli/issues/100)
- [spo customaction set](../cmd/spo/customaction/customaction-set.md) - updates a user custom action for site or site collection [#212](https://github.com/pnp/office365-cli/issues/212)

### Changes

- fixed issue with prompts in non-interactive mode [#142](https://github.com/pnp/office365-cli/issues/142)
- added information about the current user to status commands [#202](https://github.com/pnp/office365-cli/issues/202)
- fixed issue with completing input that doesn't match commands [#222](https://github.com/pnp/office365-cli/issues/222)
- fixed issue with escaping numeric input [#226](https://github.com/pnp/office365-cli/issues/226)
- changed the [aad oauth2grant list](../cmd/aad/oauth2grant/oauth2grant-list.md), [spo app list](../cmd/spo/app/app-list.md), [spo customaction list](../cmd/spo/customaction/customaction-list.md), [spo site list](../cmd/spo/site/site-list.md) commands to list all properties for output type JSON [#232](https://github.com/pnp/office365-cli/issues/232), [#233](https://github.com/pnp/office365-cli/issues/233), [#234](https://github.com/pnp/office365-cli/issues/234), [#235](https://github.com/pnp/office365-cli/issues/235)
- fixed issue with generating clink completion file [#252](https://github.com/pnp/office365-cli/issues/252)
- added [user guide](../user-guide/installing-cli.md) [#236](https://github.com/pnp/office365-cli/issues/236), [#237](https://github.com/pnp/office365-cli/issues/237), [#238](https://github.com/pnp/office365-cli/issues/238), [#239](https://github.com/pnp/office365-cli/issues/239)

## [v0.4.0](https://github.com/pnp/office365-cli/releases/tag/v0.4.0)

### Breaking changes

- renamed the `spo cdn origin set` command to [spo cdn origin add](../cmd/spo/cdn/cdn-origin-add.md) [#184](https://github.com/pnp/office365-cli/issues/184)

### New commands

**SharePoint Online:**

- [spo customaction list](../cmd/spo/customaction/customaction-list.md) - lists user custom actions for site or site collection [#19](https://github.com/pnp/office365-cli/issues/19)
- [spo site get](../cmd/spo/site/site-get.md) - gets information about the specific site collection [#114](https://github.com/pnp/office365-cli/issues/114)
- [spo site list](../cmd/spo/site/site-list.md) - lists modern sites of the given type [#115](https://github.com/pnp/office365-cli/issues/115)
- [spo site add](../cmd/spo/site/site-add.md) - creates new modern site [#116](https://github.com/pnp/office365-cli/issues/116)
- [spo app remove](../cmd/spo/app/app-remove.md) - removes the specified app from the tenant app catalog [#9](https://github.com/pnp/office365-cli/issues/9)
- [spo site appcatalog add](../cmd/spo/site/site-appcatalog-add.md) - creates a site collection app catalog in the specified site [#63](https://github.com/pnp/office365-cli/issues/63)
- [spo site appcatalog remove](../cmd/spo/site/site-appcatalog-remove.md) - removes site collection scoped app catalog from site [#64](https://github.com/pnp/office365-cli/issues/64)
- [spo serviceprincipal permissionrequest list](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-list.md) - lists pending permission requests [#152](https://github.com/pnp/office365-cli/issues/152)
- [spo serviceprincipal permissionrequest approve](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md) - approves the specified permission request [#153](https://github.com/pnp/office365-cli/issues/153)
- [spo serviceprincipal permissionrequest deny](../cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md) - denies the specified permission request [#154](https://github.com/pnp/office365-cli/issues/154)
- [spo serviceprincipal grant list](../cmd/spo/serviceprincipal/serviceprincipal-grant-list.md) - lists permissions granted to the service principal [#155](https://github.com/pnp/office365-cli/issues/155)
- [spo serviceprincipal grant revoke](../cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md) - revokes the specified set of permissions granted to the service principal [#155](https://github.com/pnp/office365-cli/issues/156)
- [spo serviceprincipal set](../cmd/spo/serviceprincipal/serviceprincipal-set.md) - enable or disable the service principal [#157](https://github.com/pnp/office365-cli/issues/157)
- [spo customaction add](../cmd/spo/customaction/customaction-add.md) - adds a user custom action for site or site collection [#18](https://github.com/pnp/office365-cli/issues/18)
- [spo externaluser list](../cmd/spo/externaluser/externaluser-list.md) - lists external users in the tenant [#27](https://github.com/pnp/office365-cli/issues/27)

**Azure Active Directory Graph:**

- [aad login](../cmd/aad/login.md) - log in to the Azure Active Directory Graph [#160](https://github.com/pnp/office365-cli/issues/160)
- [aad logout](../cmd/aad/logout.md) - log out from Azure Active Directory Graph [#161](https://github.com/pnp/office365-cli/issues/161)
- [aad status](../cmd/aad/status.md) - shows Azure Active Directory Graph login status [#162](https://github.com/pnp/office365-cli/issues/162)
- [aad sp get](../cmd/aad/sp/sp-get.md) - gets information about the specific service principal [#158](https://github.com/pnp/office365-cli/issues/158)
- [aad oauth2grant list](../cmd/aad/oauth2grant/oauth2grant-list.md) - lists OAuth2 permission grants for the specified service principal [#159](https://github.com/pnp/office365-cli/issues/159)
- [aad oauth2grant add](../cmd/aad/oauth2grant/oauth2grant-add.md) - grant the specified service principal OAuth2 permissions to the specified resource [#164](https://github.com/pnp/office365-cli/issues/164)
- [aad oauth2grant set](../cmd/aad/oauth2grant/oauth2grant-set.md) - update OAuth2 permissions for the service principal [#163](https://github.com/pnp/office365-cli/issues/163)
- [aad oauth2grant remove](../cmd/aad/oauth2grant/oauth2grant-remove.md) - remove specified service principal OAuth2 permissions [#165](https://github.com/pnp/office365-cli/issues/165)

### Changes

- added support for persisting connection [#46](https://github.com/pnp/office365-cli/issues/46)
- fixed authentication bug in `spo app install`, `spo app uninstall` and `spo app upgrade` commands when connected to the tenant admin site [#118](https://github.com/pnp/office365-cli/issues/118)
- fixed authentication bug in the `spo customaction get` command when connected to the tenant admin site [#113](https://github.com/pnp/office365-cli/issues/113)
- fixed bug in rendering help for commands when using the `--help` option [#104](https://github.com/pnp/office365-cli/issues/104)
- added detailed output to the `spo customaction get` command [#93](https://github.com/pnp/office365-cli/issues/93)
- improved collecting telemetry [#130](https://github.com/pnp/office365-cli/issues/130), [#131](https://github.com/pnp/office365-cli/issues/131), [#132](https://github.com/pnp/office365-cli/issues/132), [#133](https://github.com/pnp/office365-cli/issues/133)
- added support for the `skipFeatureDeployment` flag to the [spo app deploy](../cmd/spo/app/app-deploy.md) command [#134](https://github.com/pnp/office365-cli/issues/134)
- wrapped executing commands in `try..catch` [#109](https://github.com/pnp/office365-cli/issues/109)
- added serializing objects in log [#108](https://github.com/pnp/office365-cli/issues/108)
- added support for autocomplete in Zsh, Bash and Fish and Clink (cmder) on Windows [#141](https://github.com/pnp/office365-cli/issues/141), [#190](https://github.com/pnp/office365-cli/issues/190)

## [v0.3.0](https://github.com/pnp/office365-cli/releases/tag/v0.3.0)

### New commands

**SharePoint Online:**

- [spo customaction get](../cmd/spo/customaction/customaction-get.md) - gets information about the specific user custom action [#20](https://github.com/pnp/office365-cli/issues/20)

### Changes

- changed command output to silent [#47](https://github.com/pnp/office365-cli/issues/47)
- added user-agent string to all requests [#52](https://github.com/pnp/office365-cli/issues/52)
- refactored `spo cdn get` and `spo storageentity set` to use the `getRequestDigest` helper [#78](https://github.com/pnp/office365-cli/issues/78) and [#80](https://github.com/pnp/office365-cli/issues/80)
- added common handler for rejected OData promises [#59](https://github.com/pnp/office365-cli/issues/59)
- added Google Analytics code to documentation [#84](https://github.com/pnp/office365-cli/issues/84)
- added support for formatting command output as JSON [#48](https://github.com/pnp/office365-cli/issues/48)

## [v0.2.0](https://github.com/pnp/office365-cli/releases/tag/v0.2.0)

### New commands

**SharePoint Online:**

- [spo app add](../cmd/spo/app/app-add.md) - add an app to the specified SharePoint Online app catalog [#3](https://github.com/pnp/office365-cli/issues/3)
- [spo app deploy](../cmd/spo/app/app-deploy.md) - deploy the specified app in the tenant app catalog [#7](https://github.com/pnp/office365-cli/issues/7)
- [spo app get](../cmd/spo/app/app-get.md) - get information about the specific app from the tenant app catalog [#2](https://github.com/pnp/office365-cli/issues/2)
- [spo app install](../cmd/spo/app/app-install.md) - install an app from the tenant app catalog in the site [#4](https://github.com/pnp/office365-cli/issues/4)
- [spo app list](../cmd/spo/app/app-list.md) - list apps from the tenant app catalog [#1](https://github.com/pnp/office365-cli/issues/1)
- [spo app retract](../cmd/spo/app/app-retract.md) - retract the specified app from the tenant app catalog [#8](https://github.com/pnp/office365-cli/issues/8)
- [spo app uninstall](../cmd/spo/app/app-uninstall.md) - uninstall an app from the site [#5](https://github.com/pnp/office365-cli/issues/5)
- [spo app upgrade](../cmd/spo/app/app-upgrade.md) - upgrade app in the specified site [#6](https://github.com/pnp/office365-cli/issues/6)

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
- [spo login](../cmd/spo/login.md) - log in to a SharePoint Online site
- [spo logout](../cmd/spo/logout.md) - log out from SharePoint
- [spo status](../cmd/spo/status.md) - show SharePoint Online login status
- [spo storageentity get](../cmd/spo/storageentity/storageentity-get.md) - get value of a tenant property
- [spo storageentity list](../cmd/spo/storageentity/storageentity-list.md) - list all tenant properties
- [spo storageentity remove](../cmd/spo/storageentity/storageentity-remove.md) - remove a tenant property
- [spo storageentity set](../cmd/spo/storageentity/storageentity-set.md) - set a tenant property