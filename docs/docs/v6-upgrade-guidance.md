# v6 Upgrade Guidance

The v6 of CLI for Microsoft 365 introduces several breaking changes. To help you upgrade to the latest version of CLI for Microsoft 365, we've listed those changes along with any actions you may need to take.

## Consolidated SharePoint Online site commands

In CLI for Microsoft 365 we had several commands that were originally targeted at classic SharePoint sites. All functionality in these commands has been merged with the regular SharePoint site commands and deprecated as a result. They have therefore been dropped in v6. The commands that were dropped are:

Command | Merged with 
--|--
`spo site classic add [options]` | [spo site add](./cmd/spo/site/site-add.md)
`spo site classic list [options]` | [spo site list](./cmd/spo/site/site-list.md)
`spo site classic set [options]` | [spo site set](./cmd/spo/site/site-set.md)

To fix a resulting issue with the `spo site list` command, the default value of the `type` option of that command has been dropped. 

### What action do I need to take?

If any script references a `spo site classic` command, replace them by the respective `spo site` command. The options have not changed and the output of the commands has not changed as well. But be sure to test your scripts as there might be slight differences in how the commands behave.

## Dropped executeWithLimitedPermission option on spo list list

In version 5 of the CLI for Microsoft 365, the `spo list list` command could only be executed with site owner permissions. To update this without introducing a breaking change, we temporarily added the option `--executeWithLimitedPermission` to be able to execute it as a site visitor or member as well. In v6 we dropped this option while changing the command in such a way that you do not need to be a site owner anymore.

### What action do I need to take?

Check if you use the command in combination with the `executeWithLimitedPermission` option. The shape of the returned data has changed slightly, so be sure your script still works as intended.

## Removed deprecated aliasses

There are several commands that have aliasses that were removed because of new insights in how to structure commands in the CLI for Microsoft 365. The following aliasses were dropped:

Alias | Command | Reason
--|--|--
`aad app delete` | [aad app remove](./cmd/aad/app/app-remove.md) | Dropped to align with the naming convention.
`aad app role delete` | [aad app role remove](./cmd/aad/app/app-role-remove.md) | Dropped to align with the naming convention.
`aad o365group restore` | [aad o365group recyclebinitem restore](./cmd/aad/o365group/o365group-recyclebinitem-restore.md) | Renamed to better match intention and naming convention.
`outlook sendmail` | [outlook mail send](./cmd/outlook/mail/mail-send.md) | Renamed to better match intention and naming convention.
`planner plan details get` | [planner plan get](./cmd/planner/plan/plan-get.md) | Functionality merged in a single get-command.
`planner task details get` | [planner task get](./cmd/planner/task/task-get.md) | Functionality merged in a single get-command.
`teams conversationmember add` | [teams channel member add](./cmd/teams/channel/channel-member-add.md) | Renamed to better match intention and naming convention.
`teams conversationmember list` | [teams channel member list](./cmd/teams/channel/channel-member-list.md) | Renamed to better match intention and naming convention.
`teams conversationmember remove` | [teams channel member remove](./cmd/teams/channel/channel-member-remove.md) | Renamed to better match intention and naming convention.

### What action do I need to take?

Search for places in your scripts where you used an alias. Aliasses can be replaced with their related command without worrying about changed functionality.

## Dropped deprecated command option planName for planner commands

There are several commands in the planner workload on which we renamed the `planName` option to `planTitle` to better cover their intentions. The old option name `planName` was kept around as a deprecated option. For v6 this has been dropped. The following list of commands is affected by this change:

- [planner bucket add](./cmd/planner/bucket/bucket-add.md)
- [planner bucket get](./cmd/planner/bucket/bucket-get.md)
- [planner bucket list](./cmd/planner/bucket/bucket-list.md)
- [planner bucket remove](./cmd/planner/bucket/bucket-remove.md)
- [planner bucket set](./cmd/planner/bucket/bucket-set.md)
- [planner task add](./cmd/planner/task/task-add.md)
- [planner task get](./cmd/planner/task/task-get.md)
- [planner task list](./cmd/planner/task/task-list.md)
- [planner task set](./cmd/planner/task/task-set.md)

### What action do I need to take?

Take care that any script referencing `--planName` is updated to `--planTitle`.

## Dropped deprecated configuration key autoOpenBrowserOnLogin

The CLI for Microsoft 365 contains commands that return a link that the user should copy and open in the browser. At first, this was only available on the login command. A configuration key `autoOpenBrowserOnLogin` was introduced to allow the CLI to automaticaly open the browser, so that the user would not need to copy/paste the value. Because more commands were created that returned links, this configuration key has been renamed to `autoOpenLinksInBrowser` to be able to use it with more commands. The old deprecated key `autoOpenBrowserOnLogin` has been dropped in v6. 

### What action do I need to take?

If you have configured the `autoOpenBrowserOnLogin` key, you'll now need to configure the `autoOpenLinksInBrowser` to keep the same behavior. You can do this by running the following script:

```sh
m365 cli config set --key autoOpenLinksInBrowser --value true
```

## Aligned options with naming convention 

For version 6 of the CLI for Microsoft 365, a lot of command options have been renamed to align better with our naming convention. Our naming convention states that options that refer to the last noun in the command, don't need that noun as a prefix. for example: the option `--webUrl` for `m365 spo web list` has been renamed to `--url` as the last noun is `web`.

The list of commands that have been affected by this change is long:

- [aad approleassignment add](./cmd/aad/approleassignment/approleassignment-add.md)
- [aad approleassignment list](./cmd/aad/approleassignment/approleassignment-list.md)
- [aad approleassignment remove](./cmd/aad/approleassignment/approleassignment-remove.md)
- [aad o365group add](./cmd/aad/o365group/o365group-add.md)
- [aad o365group recyclebinitem list](./cmd/aad/o365group/o365group-recyclebinitem-list.md)
- [aad o365group teamify](./cmd/aad/o365group/o365group-teamify.md)
- [aad policy list](./cmd/aad/policy/policy-list.md)
- [aad sp get](./cmd/aad/sp/sp-get.md)
- [flow disable](./cmd/flow/flow-disable.md)
- [flow enable](./cmd/flow/flow-enable.md)
- [flow export](./cmd/flow/flow-export.md)
- [flow get](./cmd/flow/flow-get.md)
- [flow list](./cmd/flow/flow-list.md)
- [flow remove ](./cmd/flow/flow-remove.md)
- [flow run cancel](./cmd/flow/run/run-cancel.md)
- [flow run get](./cmd/flow/run/run-get.md)
- [flow run list](./cmd/flow/run/run-list.md)
- [flow run resubmit](./cmd/flow/run/run-resubmit.md)
- [outlook message mmove](./cmd/outlook/message/message-move.md)
- [pa connector export](./cmd/pa/connector/connector-export.md)
- [pa connector list](./cmd/pa/connector/connector-list.md)
- [pa solution reference-add](./cmd/pa/solution/solution-reference-add.md)
- [spfx package generate](./cmd/spfx/package/package-generate.md)
- [spo app add](./cmd/spo/app/app-add.md)
- [spo app deploy](./cmd/spo/app/app-deploy.md)
- [spo app get](./cmd/spo/app/app-get.md)
- [spo app install](./cmd/spo/app/app-install.md)
- [spo app list](./cmd/spo/app/app-list.md)
- [spo app remove](./cmd/spo/app/app-remove.md)
- [spo app retract](./cmd/spo/app/app-retract.md)
- [spo app uninstall](./cmd/spo/app/app-uninstall.md)
- [spo app upgrade](./cmd/spo/app/app-upgrade.md)
- [spo apppage set](./cmd/spo/apppage/apppage-set.md)
- [spo cdn policy list](./cmd/spo/cdn/cdn-policy-list.md)
- [spo cdn policy set](./cmd/spo/cdn/cdn-policy-set.md)
- [spo contenttype field set](./cmd/spo/contenttype/contenttype-field-set.md)
- [spo customaction add](./cmd/spo/customaction/customaction-add.md)
- [spo customaction clear](./cmd/spo/customaction/customaction-clear.md)
- [spo customaction get](./cmd/spo/customaction/customaction-get.md)
- [spo customaction list](./cmd/spo/customaction/customaction-list.md)
- [spo customaction remove](./cmd/spo/customaction/customaction-remove.md)
- [spo customaction set](./cmd/spo/customaction/customaction-set.md)
- [spo feature disable](./cmd/spo/feature/feature-disable.md)
- [spo feature enable](./cmd/spo/feature/feature-enable.md)
- [spo feature list](./cmd/spo/feature/feature-list.md)
- [spo field get](./cmd/spo/field/field-get.md)
- [spo field remove](./cmd/spo/field/field-remove.md)
- [spo file checkin](./cmd/spo/file/file-checkin.md)
- [spo file checkout](./cmd/spo/file/file-checkout.md)
- [spo file sharinginfo get](./cmd/spo/file/file-sharinginfo-get.md)
- [spo folder get](./cmd/spo/folder/folder-get.md)
- [spo folder remove](./cmd/spo/folder/folder-remove.md)
- [spo folder rename](./cmd/spo/folder/folder-rename.md)
- [spo hubsite connect](./cmd/spo/hubsite/hubsite-connect.md)
- [spo hubsite disconnect](./cmd/spo/hubsite/hubsite-disconnect.md)
- [spo hubsite register](./cmd/spo/hubsite/hubsite-register.md)
- [spo hubsite rights grant](./cmd/spo/hubsite/hubsite-rights-grant.md)
- [spo hubsite rights revoke](./cmd/spo/hubsite/hubsite-rights-revoke.md)
- [spo knowledgehub set](./cmd/spo/knowledgehub/knowledgehub-set.md)
- [spo list contenttype add](./cmd/spo/list/list-contenttype-add.md)
- [spo list contenttype remove](./cmd/spo/list/list-contenttype-remove.md)
- [spo list view field add](./cmd/spo/list/list-view-field-add.md)
- [spo list view field remove](./cmd/spo/list/list-view-field-remove.md)
- [spo list view field set](./cmd/spo/list/list-view-field-set.md)
- [spo list view get](./cmd/spo/list/list-view-get.md)
- [spo list view remove](./cmd/spo/list/list-view-remove.md)
- [spo list view set](./cmd/spo/list/list-view-set.md)
- [spo listitem record declare](./cmd/spo/listitem/listitem-record-declare.md)
- [spo listitem record undeclare](./cmd/spo/listitem/listitem-record-undeclare.md)
- [spo page column get](./cmd/spo/page/page-column-get.md)
- [spo page column list](./cmd/spo/page/page-column-list.md)
- [spo page control get](./cmd/spo/page/page-control-get.md)
- [spo page control list](./cmd/spo/page/page-control-list.md)
- [spo page control set](./cmd/spo/page/page-control-set.md)
- [spo page section add](./cmd/spo/page/page-section-add.md)
- [spo page section get](./cmd/spo/page/page-section-get.md)
- [spo page section list](./cmd/spo/page/page-section-list.md)
- [spo serviceprincipal grant revoke](./cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)
- [spo serviceprincipal permissionrequest approve](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)
- [spo serviceprincipal permissionrequest deny](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)
- [spo site appcatalog add](./cmd/spo/site/site-appcatalog-add.md)
- [spo site appcatalog remove](./cmd/spo/site/site-appcatalog-remove.md)
- [spo site apppermission get](./cmd/spo/site/site-apppermission-get.md)
- [spo site apppermission remove](./cmd/spo/site/site-apppermission-remove.md)
- [spo site apppermission set](./cmd/spo/site/site-apppermission-set.md)
- [spo site chrome set](./cmd/spo/site/site-chrome-set.md)
- [spo site groupify](./cmd/spo/site/site-groupify.md)
- [spo site rename](./cmd/spo/site/site-rename.md)
- [spo sitedesign rights grant](./cmd/spo/sitedesign/sitedesign-rights-grant.md)
- [spo sitedesign rights list](./cmd/spo/sitedesign/sitedesign-rights-list.md)
- [spo sitedesign rights revoke](./cmd/spo/sitedesign/sitedesign-rights-revoke.md)
- [spo sitedesign task get](./cmd/spo/sitedesign/sitedesign-task-get.md)
- [spo sitedesign task remove](./cmd/spo/sitedesign/sitedesign-task-remove.md)
- [spo tenant recyclebinitem remove](./cmd/spo/tenant/tenant-recyclebinitem-remove.md)
- [spo tenant recyclebinitem restore](./cmd/spo/tenant/tenant-recyclebinitem-restore.md)
- [spo web add](./cmd/spo/web/web-add.md)
- [spo web get](./cmd/spo/web/web-get.md)
- [spo web list](./cmd/spo/web/web-list.md)
- [spo web reindex](./cmd/spo/web/web-reindex.md)
- [spo web remove](./cmd/spo/web/web-remove.md)
- [spo web set](./cmd/spo/web/web-set.md)
- [teams app app install](./cmd/teams/app/app-install.md)
- [teams app app uninstall](./cmd/teams/app/app-uninstall.md)
- [teams channel get](./cmd/teams/channel/channel-get.md)
- [teams channel remove](./cmd/teams/channel/channel-remove.md)
- [teams channel set](./cmd/teams/channel/channel-set.md)
- [teams message get](./cmd/teams/message/message-get.md)
- [teams tab get](./cmd/teams/tab/tab-get.md)
- [teams tab remove](./cmd/teams/tab/tab-remove.md)
- [teams team archive](./cmd/teams/team/team-archive.md)
- [teams team clone](./cmd/teams/team/team-clone.md)
- [teams team remove](./cmd/teams/team/team-remove.md)
- [teams team set](./cmd/teams/team/team-set.md)
- [teams team unarchive](./cmd/teams/team/team-unarchive.md)
- [teams user app add](./cmd/teams/user/user-app-add.md)
- [teams user app remove](./cmd/teams/user/user-app-remove.md)
- [viva connections app create](./cmd/viva/connections/connections-app-create.md)
- [yammer group user add](./cmd/yammer/group/group-user-add.md)
- [yammer group user remove](./cmd/yammer/group/group-user-remove.md)
- [yammer message like set](./cmd/yammer/message/message-like-set.md)
- [yammer user get](./cmd/yammer/user/user-get.md)

### What action do I need to take?

When running existing scripts after an update to v6, verify that your command options are correctly written. If in doubt, employ our great [shell completion functionality](./user-guide/completion.md) to quickly check what an option name should be. 
