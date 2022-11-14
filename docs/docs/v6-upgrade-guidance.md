# v6 Upgrade Guidance

The v6 of CLI for Microsoft 365 introduces several breaking changes. To help you upgrade to the latest version of CLI for Microsoft 365, we've listed those changes along with any actions you may need to take.

## Consolidated SharePoint Online site commands

In CLI for Microsoft 365 we had several commands that were originally targeted at classic SharePoint sites. All functionality in these commands has been merged with the regular SharePoint site commands and deprecated as a result. They have therefore been removed in v6. The commands that were removed are:

Command | Merged with
--|--
`spo site classic add` | [spo site add](./cmd/spo/site/site-add.md)
`spo site classic list` | [spo site list](./cmd/spo/site/site-list.md)
`spo site classic set` | [spo site set](./cmd/spo/site/site-set.md)

To fix a resulting issue with the `spo site list` command, the default value of the `type` option of that command has been removed.

### What action do I need to take?

Replace references to `spo site classic *` commands, with the respective `spo site *` command. The options have not changed and the output of the commands has not changed as well. After updating the references test your scripts as there might be slight differences in how the commands behave.

## Removed the `executeWithLimitedPermission` option in the `spo list list` command

In version 5 of the CLI for Microsoft 365, the `spo list list` command could only be executed with site owner permissions. To update this without introducing a breaking change, we temporarily added the `--executeWithLimitedPermission` option to be able to execute it as a site visitor or member as well. In v6 we removed this option while changing the command in such a way that you do not need to be a site owner anymore.

### What action do I need to take?

Remove the references to the `executeWithLimitedPermission` option in your scripts. Verify that your scripts work as intended with the new data structure returned by the `spo list list` command.

## Removed deprecated aliases

We removed several aliases to align commands with our naming convention. The following aliases were removed:

Alias | Command | Reason
--|--|--
`aad app delete` | [aad app remove](./cmd/aad/app/app-remove.md) | Removed to align with the naming convention.
`aad app role delete` | [aad app role remove](./cmd/aad/app/app-role-remove.md) | Removed to align with the naming convention.
`aad o365group restore` | [aad o365group recyclebinitem restore](./cmd/aad/o365group/o365group-recyclebinitem-restore.md) | Renamed to better match intention and naming convention.
`outlook sendmail` | [outlook mail send](./cmd/outlook/mail/mail-send.md) | Renamed to better match intention and naming convention.
`planner plan details get` | [planner plan get](./cmd/planner/plan/plan-get.md) | Functionality merged in a single get-command.
`planner task details get` | [planner task get](./cmd/planner/task/task-get.md) | Functionality merged in a single get-command.
`teams conversationmember add` | [teams channel member add](./cmd/teams/channel/channel-member-add.md) | Renamed to better match intention and naming convention.
`teams conversationmember list` | [teams channel member list](./cmd/teams/channel/channel-member-list.md) | Renamed to better match intention and naming convention.
`teams conversationmember remove` | [teams channel member remove](./cmd/teams/channel/channel-member-remove.md) | Renamed to better match intention and naming convention.
`spo hubsite theme sync` | [spo site hubsite theme sync](./cmd/spo/site/site-hubsite-theme-sync.md) | Renamed to better match intention and naming convention.
`spo hubsite connect` | [spo site hubsite connect](./cmd/spo/site/site-hubsite-connect.md) | Renamed to better match intention and naming convention.
`spo hubsite disconnect` | [spo site hubsite disconnect](./cmd/spo/site/site-hubsite-disconnect.md) | Renamed to better match intention and naming convention.

### What action do I need to take?

Replace any of the aliases mentioned above with the corresponding command name. The functionality of the command hasn't changed.

## In `planner` commands, removed the deprecated `planName` option

In several planner commands we renamed the `planName` option to `planTitle` to align with the underlying API. Along with introducing the `planTitle` option, we deprecated the old `planName` option. In v6 of the CLI we removed the `planName` option. The following list of commands is affected by this change:

- [planner bucket add](./cmd/planner/bucket/bucket-add.md)
- [planner bucket get](./cmd/planner/bucket/bucket-get.md)
- [planner bucket list](./cmd/planner/bucket/bucket-list.md)
- [planner bucket remove](./cmd/planner/bucket/bucket-remove.md)
- [planner bucket set](./cmd/planner/bucket/bucket-set.md)
- [planner plan get](./cmd/planner/plan/plan-get.md)
- [planner task add](./cmd/planner/task/task-add.md)
- [planner task get](./cmd/planner/task/task-get.md)
- [planner task list](./cmd/planner/task/task-list.md)
- [planner task set](./cmd/planner/task/task-set.md)

### What action do I need to take?

Replace the reference to the `--planName` option with `--planTitle`.

## Removed the deprecated `autoOpenBrowserOnLogin` configuration key

The CLI for Microsoft 365 contains commands that return a link that the user should copy and open in the browser. At first, this was only available in the `login` command. We introduced the `autoOpenBrowserOnLogin` configuration key was introduced to have the CLI automatically open the browser, so that you don't have to copy/paste the URL manually. As we introduced this functionality to other commands, we renamed this configuration key to `autoOpenLinksInBrowser`. In v6 of the CLI, we removed the deprecated `autoOpenBrowserOnLogin` key.

### What action do I need to take?

If you have configured the `autoOpenBrowserOnLogin` key, you'll now need to configure the `autoOpenLinksInBrowser` key to keep the same behavior. You can do this by running the following script:

```sh
m365 cli config set --key autoOpenLinksInBrowser --value true
```

## In `teams channel` commands, changed short options

In the following commands we've changed some shorts:

- [teams channel get](./cmd/teams/channel/channel-get.md)
- [teams channel set](./cmd/teams/channel/channel-set.md)
- [teams channel remove](./cmd/teams/channel/channel-remove.md)

The following shorts where changed:

- Where we used `-c, --id`, we changed it to `-i, --id`.
- Where we used `-i, --teamId`, we changed it to `--teamId`.

### What action do I need to take?

Update the reference to the short options in your scripts.

## Updated `teams app publish` command output

In the past versions, `teams app publish` returned just the app ID of the published app, or nothing at all. This has been adjusted, now the command will return the entire result object.

Previous JSON command output:
```json
"fbdfd207-83ee-45d8-9c98-5039a1a01207"
```

Current JSON command output:
```json
{
    "id": "fbdfd207-83ee-45d8-9c98-5039a1a01207",
    "externalId": "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
    "displayName": "Test App",
    "distributionMethod": "organization"
}
```

### What action do I need to take?

Update your scripts to read the `id` property of the command output.

## Updated `spo eventreceiver get` command output

In the past versions, `spo eventreceiver get` returned an array with a single object. This has been adjusted, now the command will return the object.

Previous JSON command output:
```json
[
  {
    "ReceiverAssembly": "Microsoft.Office.Server.UserProfiles, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
    "ReceiverClass": "Microsoft.Office.Server.UserProfiles.ContentFollowingWebEventReceiver",
    "ReceiverId": "c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec",
    "ReceiverName": "PnP Test Receiver",
    "SequenceNumber": 10000,
    "Synchronization": 2,
    "EventType": 10204,
    "ReceiverUrl": "https://northeurope1-0.pushnp.svc.ms/notifications?token=b4c0def2-a5ea-490a-bb85-c2e423b1384b"
  }
]
```

Current JSON command output:
```json
{
  "ReceiverAssembly": "Microsoft.Office.Server.UserProfiles, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
  "ReceiverClass": "Microsoft.Office.Server.UserProfiles.ContentFollowingWebEventReceiver",
  "ReceiverId": "c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec",
  "ReceiverName": "PnP Test Receiver",
  "SequenceNumber": 10000,
  "Synchronization": 2,
  "EventType": 10204,
  "ReceiverUrl": "https://northeurope1-0.pushnp.svc.ms/notifications?token=b4c0def2-a5ea-490a-bb85-c2e423b1384b"
}
```

### What action do I need to take?

Update your scripts to expect an object instead of an array.

## Aligned options with naming convention

As we've been adding more commands to the CLI, we noticed that several commands were using inconsistent options names. Our naming convention states that options that refer to the last noun in the command, don't need that noun as a prefix. for example: the option `--webUrl` for `m365 spo web list` has been renamed to `--url` as the last noun is `web`. In version 6 of the CLI for Microsoft 365, we updated all these options to be consistent and make it easier for you to use the CLI.

We've updated the following commands and options:

Command|Old option|New option
--|--|--
[aad approleassignment add](./cmd/aad/approleassignment/approleassignment-add.md)|`objectId`|`appObjectId`
[aad approleassignment add](./cmd/aad/approleassignment/approleassignment-add.md)|`displayName`|`appDisplayName`
[aad approleassignment list](./cmd/aad/approleassignment/approleassignment-list.md)|`objectId`|`appObjectId`
[aad approleassignment list](./cmd/aad/approleassignment/approleassignment-list.md)|`displayName`|`appDisplayName`
[aad approleassignment remove](./cmd/aad/approleassignment/approleassignment-remove.md)|`objectId`|`appObjectId`
[aad approleassignment remove](./cmd/aad/approleassignment/approleassignment-remove.md)|`displayName`|`appDisplayName`
[aad o365group add](./cmd/aad/o365group/o365group-add.md)|`isPrivate [isPrivate]`|`isPrivate` (changed to flag)
[aad o365group recyclebinitem list](./cmd/aad/o365group/o365group-recyclebinitem-list.md)|`displayName`|`groupDisplayName`
[aad o365group recyclebinitem list](./cmd/aad/o365group/o365group-recyclebinitem-list.md)|`mailNickname`|`groupMailNickname`
[aad o365group teamify](./cmd/aad/o365group/o365group-teamify.md)|`groupId`|`id`
[aad policy list](./cmd/aad/policy/policy-list.md)|`policyType`|`type`
[aad sp get](./cmd/aad/sp/sp-get.md)|`displayName`|`appDisplayName`
[aad sp get](./cmd/aad/sp/sp-get.md)|`objectId`|`appObjectId`
[flow disable](./cmd/flow/flow-disable.md)|`environment`|`environmentName`
[flow enable](./cmd/flow/flow-enable.md)|`environment`|`environmentName`
[flow export](./cmd/flow/flow-export.md)|`environment`|`environmentName`
[flow get](./cmd/flow/flow-get.md)|`environment`|`environmentName`
[flow list](./cmd/flow/flow-list.md)|`environment`|`environmentName`
[flow remove](./cmd/flow/flow-remove.md)|`environment`|`environmentName`
[flow run cancel](./cmd/flow/run/run-cancel.md)|`flow`|`flowName`
[flow run cancel](./cmd/flow/run/run-cancel.md)|`environment`|`environmentName`
[flow run get](./cmd/flow/run/run-get.md)|`flow`|`flowName`
[flow run get](./cmd/flow/run/run-get.md)|`environment`|`environmentName`
[flow run list](./cmd/flow/run/run-list.md)|`flow`|`flowName`
[flow run list](./cmd/flow/run/run-list.md)|`environment`|`environmentName`
[flow run resubmit](./cmd/flow/run/run-resubmit.md)|`flow`|`flowName`
[flow run resubmit](./cmd/flow/run/run-resubmit.md)|`environment`|`environmentName`
[outlook message move](./cmd/outlook/message/message-move.md)|`messageId`|`id`
[pa connector export](./cmd/pa/connector/connector-export.md)|`environment`|`environmentName`
[pa connector list](./cmd/pa/connector/connector-list.md)|`environment`|`environmentName`
[pa solution reference add](./cmd/pa/solution/solution-reference-add.md)|`path`|`projectPath`
[spfx package generate](./cmd/spfx/package/package-generate.md)|`packageName`|`name`
[spo app add](./cmd/spo/app/app-add.md)|`scope`|`appCatalogScope`
[spo app deploy](./cmd/spo/app/app-deploy.md)|`scope`|`appCatalogScope`
[spo app get](./cmd/spo/app/app-get.md)|`scope`|`appCatalogScope`
[spo app install](./cmd/spo/app/app-install.md)|`scope`|`appCatalogScope`
[spo app list](./cmd/spo/app/app-list.md)|`scope`|`appCatalogScope`
[spo app remove](./cmd/spo/app/app-remove.md)|`scope`|`appCatalogScope`
[spo app retract](./cmd/spo/app/app-retract.md)|`scope`|`appCatalogScope`
[spo app uninstall](./cmd/spo/app/app-uninstall.md)|`scope`|`appCatalogScope`
[spo app upgrade](./cmd/spo/app/app-upgrade.md)|`scope`|`appCatalogScope`
[spo apppage set](./cmd/spo/apppage/apppage-set.md)|`pageName`|`name`
[spo cdn policy list](./cmd/spo/cdn/cdn-policy-list.md)|`type`|`cdnType`
[spo cdn policy set](./cmd/spo/cdn/cdn-policy-set.md)|`type`|`cdnType`
[spo contenttype field set](./cmd/spo/contenttype/contenttype-field-set.md)|`fieldId`|`id`
[spo customaction add](./cmd/spo/customaction/customaction-add.md)|`url`|`webUrl`
[spo customaction clear](./cmd/spo/customaction/customaction-clear.md)|`url`|`webUrl`
[spo customaction get](./cmd/spo/customaction/customaction-get.md)|`url`|`webUrl`
[spo customaction list](./cmd/spo/customaction/customaction-list.md)|`url`|`webUrl`
[spo customaction remove](./cmd/spo/customaction/customaction-remove.md)|`url`|`webUrl`
[spo customaction set](./cmd/spo/customaction/customaction-set.md)|`url`|`webUrl`
[spo feature disable](./cmd/spo/feature/feature-disable.md)|`featureId`|`id`
[spo feature disable](./cmd/spo/feature/feature-disable.md)|`url`|`webUrl`
[spo feature enable](./cmd/spo/feature/feature-enable.md)|`featureId`|`id`
[spo feature enable](./cmd/spo/feature/feature-enable.md)|`url`|`webUrl`
[spo feature list](./cmd/spo/feature/feature-list.md)|`url`|`webUrl`
[spo field get](./cmd/spo/field/field-get.md)|`fieldTitle`|`title`
[spo field remove](./cmd/spo/field/field-remove.md)|`fieldTitle`|`title`
[spo file checkin](./cmd/spo/file/file-checkin.md)|`fileUrl`|`url`
[spo file checkout](./cmd/spo/file/file-checkout.md)|`fileUrl`|`url`
[spo file sharinginfo get](./cmd/spo/file/file-sharinginfo-get.md)|`url`|`fileUrl`
[spo file sharinginfo get](./cmd/spo/file/file-sharinginfo-get.md)|`id`|`fileId`
[spo folder get](./cmd/spo/folder/folder-get.md)|`folderUrl`|`url`
[spo folder remove](./cmd/spo/folder/folder-remove.md)|`folderUrl`|`url`
[spo folder rename](./cmd/spo/folder/folder-rename.md)|`folderUrl`|`url`
[spo site hubsite connect](./cmd/spo/site/site-hubsite-connect.md)|`url`|`siteUrl`
[spo site hubsite disconnect](./cmd/spo/site/site-hubsite-disconnect.md)|`url`|`siteUrl`
[spo hubsite register](./cmd/spo/hubsite/hubsite-register.md)|`url`|`siteUrl`
[spo hubsite rights grant](./cmd/spo/hubsite/hubsite-rights-grant.md)|`url`|`hubSiteUrl`
[spo hubsite rights revoke](./cmd/spo/hubsite/hubsite-rights-revoke.md)|`url`|`hubSiteUrl`
[spo knowledgehub set](./cmd/spo/knowledgehub/knowledgehub-set.md)|`url`|`siteUrl`
[spo list contenttype add](./cmd/spo/list/list-contenttype-add.md)|`contentTypeId`|`id`
[spo list contenttype remove](./cmd/spo/list/list-contenttype-remove.md)|`contentTypeId`|`id`
[spo list view field add](./cmd/spo/list/list-view-field-add.md)|`fieldId`|`id`
[spo list view field add](./cmd/spo/list/list-view-field-add.md)|`fieldTitle`|`title`
[spo list view field add](./cmd/spo/list/list-view-field-add.md)|`fieldPosition`|`position`
[spo list view field remove](./cmd/spo/list/list-view-field-remove.md)|`fieldId`|`id`
[spo list view field remove](./cmd/spo/list/list-view-field-remove.md)|`fieldTitle`|`title`
[spo list view field set](./cmd/spo/list/list-view-field-set.md)|`fieldId`|`id`
[spo list view field set](./cmd/spo/list/list-view-field-set.md)|`fieldTitle`|`title`
[spo list view field set](./cmd/spo/list/list-view-field-set.md)|`fieldPosition`|`position`
[spo list view get](./cmd/spo/list/list-view-get.md)|`viewId`|`id`
[spo list view get](./cmd/spo/list/list-view-get.md)|`viewTitle`|`title`
[spo list view remove](./cmd/spo/list/list-view-remove.md)|`viewId`|`id`
[spo list view remove](./cmd/spo/list/list-view-remove.md)|`viewTitle`|`title`
[spo list view set](./cmd/spo/list/list-view-set.md)|`viewId`|`id`
[spo list view set](./cmd/spo/list/list-view-set.md)|`viewTitle`|`title`
[spo listitem list](./cmd/spo/listitem/listitem-list.md)|`id`|`listId`
[spo listitem list](./cmd/spo/listitem/listitem-list.md)|`title`|`listTitle`
[spo listitem record declare](./cmd/spo/listitem/listitem-record-declare.md)|`id`|`listItemId`
[spo listitem record undeclare](./cmd/spo/listitem/listitem-record-undeclare.md)|`id`|`listItemId`
[spo page column get](./cmd/spo/page/page-column-get.md)|`name`|`pageName`
[spo page column list](./cmd/spo/page/page-column-list.md)|`name`|`pageName`
[spo page control get](./cmd/spo/page/page-control-get.md)|`name`|`pageName`
[spo page control list](./cmd/spo/page/page-control-list.md)|`name`|`pageName`
[spo page control set](./cmd/spo/page/page-control-set.md)|`name`|`pageName`
[spo page section add](./cmd/spo/page/page-section-add.md)|`name`|`pageName`
[spo page section get](./cmd/spo/page/page-section-get.md)|`name`|`pageName`
[spo page section list](./cmd/spo/page/page-section-list.md)|`name`|`pageName`
[spo serviceprincipal grant revoke](./cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.md)|`grantId`|`id`
[spo serviceprincipal permissionrequest approve](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.md)|`requestId`|`id`
[spo serviceprincipal permissionrequest deny](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.md)|`requestId`|`id`
[spo site appcatalog add](./cmd/spo/site/site-appcatalog-add.md)|`url`|`siteUrl`
[spo site appcatalog remove](./cmd/spo/site/site-appcatalog-remove.md)|`url`|`siteUrl`
[spo site apppermission get](./cmd/spo/site/site-apppermission-get.md)|`permissionId`|`id`
[spo site apppermission remove](./cmd/spo/site/site-apppermission-remove.md)|`permissionId`|`id`
[spo site apppermission set](./cmd/spo/site/site-apppermission-set.md)|`permissionId`|`id`
[spo site chrome set](./cmd/spo/site/site-chrome-set.md)|`url`|`siteUrl`
[spo site groupify](./cmd/spo/site/site-groupify.md)|`siteUrl`|`url`
[spo site rename](./cmd/spo/site/site-rename.md)|`siteUrl`|`url`
[spo sitedesign rights grant](./cmd/spo/sitedesign/sitedesign-rights-grant.md)|`id`|`siteDesignId`
[spo sitedesign rights list](./cmd/spo/sitedesign/sitedesign-rights-list.md)|`id`|`siteDesignId`
[spo sitedesign rights revoke](./cmd/spo/sitedesign/sitedesign-rights-revoke.md)|`id`|`siteDesignId`
[spo sitedesign task get](./cmd/spo/sitedesign/sitedesign-task-get.md)|`taskId`|`id`
[spo sitedesign task remove](./cmd/spo/sitedesign/sitedesign-task-remove.md)|`taskId`|`id`
[spo tenant recyclebinitem remove](./cmd/spo/tenant/tenant-recyclebinitem-remove.md)|`url`|`siteUrl`
[spo tenant recyclebinitem restore](./cmd/spo/tenant/tenant-recyclebinitem-restore.md)|`url`|`siteUrl`
[spo web add](./cmd/spo/web/web-add.md)|`webUrl`|`url`
[spo web get](./cmd/spo/web/web-get.md)|`webUrl`|`url`
[spo web list](./cmd/spo/web/web-list.md)|`webUrl`|`url`
[spo web reindex](./cmd/spo/web/web-reindex.md)|`webUrl`|`url`
[spo web remove](./cmd/spo/web/web-remove.md)|`webUrl`|`url`
[spo web set](./cmd/spo/web/web-set.md)|`webUrl`|`url`
[teams app app install](./cmd/teams/app/app-install.md)|`appId`|`id`
[teams app app uninstall](./cmd/teams/app/app-uninstall.md)|`appId`|`id`
[teams channel get](./cmd/teams/channel/channel-get.md)|`channelId`|`id`
[teams channel get](./cmd/teams/channel/channel-get.md)|`channelName`|`name`
[teams channel remove](./cmd/teams/channel/channel-remove.md)|`channelId`|`id`
[teams channel remove](./cmd/teams/channel/channel-remove.md)|`channelName`|`name`
[teams channel set](./cmd/teams/channel/channel-set.md)|`channelName`|`name`
[teams message get](./cmd/teams/message/message-get.md)|`messageId`|`id`
[teams tab get](./cmd/teams/tab/tab-get.md)|`tabId`|`id`
[teams tab get](./cmd/teams/tab/tab-get.md)|`tabName`|`name`
[teams tab remove](./cmd/teams/tab/tab-remove.md)|`tabId`|`id`
[teams team archive](./cmd/teams/team/team-archive.md)|`teamId`|`id`
[teams team clone](./cmd/teams/team/team-clone.md)|`teamId`|`id`
[teams team remove](./cmd/teams/team/team-remove.md)|`teamId`|`id`
[teams team set](./cmd/teams/team/team-set.md)|`teamId`|`id`
[teams team unarchive](./cmd/teams/team/team-unarchive.md)|`teamId`|`id`
[teams user app add](./cmd/teams/user/user-app-add.md)|`appId`|`id`
[teams user app remove](./cmd/teams/user/user-app-remove.md)|`appId`|`id`
[viva connections app create](./cmd/viva/connections/connections-app-create.md)|`appName`|`name`
[yammer group user add](./cmd/yammer/group/group-user-add.md)|`id`|`groupId`
[yammer group user add](./cmd/yammer/group/group-user-add.md)|`userId`|`id`
[yammer group user remove](./cmd/yammer/group/group-user-remove.md)|`id`|`groupId`
[yammer group user remove](./cmd/yammer/group/group-user-remove.md)|`userId`|`id`
[yammer message like set](./cmd/yammer/message/message-like-set.md)|`id`|`messageId`
[yammer user get](./cmd/yammer/user/user-get.md)|`userId`|`id`

### What action do I need to take?

If you use any of the commands listed above, ensure that you use the new option names.
