# v6 Upgrade Guidance

The v6 of CLI for Microsoft 365 introduces several breaking changes. To help you upgrade to the latest version of CLI for Microsoft 365, we've listed those changes along with any actions you may need to take.

## Consolidated SharePoint Online site commands

In CLI for Microsoft 365 we had several commands that were originally targeted at classic SharePoint sites. All functionality in these commands has been merged with the regular SharePoint site commands and deprecated as a result. They have therefore been removed in v6. The commands that were removed are:

Command | Merged with
--|--
`spo site classic add` | [spo site add](./cmd/spo/site/site-add.mdx)
`spo site classic list` | [spo site list](./cmd/spo/site/site-list.mdx)
`spo site classic set` | [spo site set](./cmd/spo/site/site-set.mdx)

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
`aad app delete` | [aad app remove](./cmd/aad/app/app-remove.mdx) | Removed to align with the naming convention.
`aad app role delete` | [aad app role remove](./cmd/aad/app/app-role-remove.mdx) | Removed to align with the naming convention.
`aad o365group restore` | [aad o365group recyclebinitem restore](./cmd/aad/o365group/o365group-recyclebinitem-restore.mdx) | Renamed to better match intention and naming convention.
`outlook sendmail` | [outlook mail send](./cmd/outlook/mail/mail-send.mdx) | Renamed to better match intention and naming convention.
`planner plan details get` | [planner plan get](./cmd/planner/plan/plan-get.mdx) | Functionality merged in a single get-command.
`planner task details get` | [planner task get](./cmd/planner/task/task-get.mdx) | Functionality merged in a single get-command.
`teams conversationmember add` | [teams channel member add](./cmd/teams/channel/channel-member-add.mdx) | Renamed to better match intention and naming convention.
`teams conversationmember list` | [teams channel member list](./cmd/teams/channel/channel-member-list.mdx) | Renamed to better match intention and naming convention.
`teams conversationmember remove` | [teams channel member remove](./cmd/teams/channel/channel-member-remove.mdx) | Renamed to better match intention and naming convention.
`spo hubsite theme sync` | [spo site hubsite theme sync](./cmd/spo/site/site-hubsite-theme-sync.mdx) | Renamed to better match intention and naming convention.
`spo hubsite connect` | [spo site hubsite connect](./cmd/spo/site/site-hubsite-connect.mdx) | Renamed to better match intention and naming convention.
`spo hubsite disconnect` | [spo site hubsite disconnect](./cmd/spo/site/site-hubsite-disconnect.mdx) | Renamed to better match intention and naming convention.

### What action do I need to take?

Replace any of the aliases mentioned above with the corresponding command name. The functionality of the command hasn't changed.

## In `planner` commands, removed the deprecated `planName` option

In several planner commands we renamed the `planName` option to `planTitle` to align with the underlying API. Along with introducing the `planTitle` option, we deprecated the old `planName` option. In v6 of the CLI we removed the `planName` option. The following list of commands is affected by this change:

- [planner bucket add](./cmd/planner/bucket/bucket-add.mdx)
- [planner bucket get](./cmd/planner/bucket/bucket-get.mdx)
- [planner bucket list](./cmd/planner/bucket/bucket-list.mdx)
- [planner bucket remove](./cmd/planner/bucket/bucket-remove.mdx)
- [planner bucket set](./cmd/planner/bucket/bucket-set.mdx)
- [planner plan get](./cmd/planner/plan/plan-get.mdx)
- [planner task add](./cmd/planner/task/task-add.mdx)
- [planner task get](./cmd/planner/task/task-get.mdx)
- [planner task list](./cmd/planner/task/task-list.mdx)
- [planner task set](./cmd/planner/task/task-set.mdx)

### What action do I need to take?

Replace the reference to the `--planName` option with `--planTitle`.

## Removed the deprecated `autoOpenBrowserOnLogin` configuration key

The CLI for Microsoft 365 contains commands that return a link that the user should copy and open in the browser. At first, this was only available in the `login` command. We introduced the `autoOpenBrowserOnLogin` configuration key was introduced to have the CLI automatically open the browser, so that you don't have to copy/paste the URL manually. As we introduced this functionality to other commands, we renamed this configuration key to `autoOpenLinksInBrowser`. In v6 of the CLI, we removed the deprecated `autoOpenBrowserOnLogin` key.

### What action do I need to take?

If you have configured the `autoOpenBrowserOnLogin` key, you'll now need to configure the `autoOpenLinksInBrowser` key to keep the same behavior. You can do this by running the following script:

```sh
m365 cli config set --key autoOpenLinksInBrowser --value true
```

## Updated `spo file copy` options

We updated the [spo file copy](./cmd/spo/file/file-copy.mdx) command. The improved functionality support copying files larger than 2GB and specify the name for the copied file. To support these changes, we had to do several changes to the command's options. When you specify an URL for options `webUrl`, `sourceUrl` and `targetUrl`, make sure that you specify a decoded URL. Specifying an encoded URL will result in a `File Not Found` error. For example, `/sites/IT/Shared%20Documents/Document.pdf` will not work while `/sites/IT/Shared Documents/Document.pdf` will work just fine.

Because of this rework, we were able to add new options, but we also removed existing ones.

**Removed options:**

- `--deleteIfAlreadyExists`
- `--allowSchemaMismatch`

**New options:**

Option | Description
--- | ---
`--nameConflictBehavior [nameConflictBehavior]` | Behavior when a document with the same name is already present at the destination. Possible values: `fail`, `replace`, `rename`. Default is `fail`.
`--newName [newName]` | New name of the destination file.
`--bypassSharedLock` | This indicates whether a file with a share lock can still be copied. Use this option to copy a file that is locked.

### What action do I need to take?

Update your scripts with the following:

- Ensure all the URLs you provide are **decoded**.
- Remove the option `--allowSchemaMismatch`.
- Replace option `--deleteIfAlreadyExists` with `--nameConflictBehavior replace`.

## In `teams channel` commands, changed short options

In the following commands we've changed some shorts:

- [teams channel get](./cmd/teams/channel/channel-get.mdx)
- [teams channel set](./cmd/teams/channel/channel-set.mdx)
- [teams channel remove](./cmd/teams/channel/channel-remove.mdx)

The following shorts where changed:

- Where we used `-c, --id`, we changed it to `-i, --id`.
- Where we used `-i, --teamId`, we changed it to `--teamId`.

### What action do I need to take?

Update the reference to the short options in your scripts.

## Updated `teams app publish` command output

In the past versions, `teams app publish` returned just the app ID of the published app, or nothing at all. This has been adjusted, now the command will return the entire result object.

v5 JSON command output:

```json
"fbdfd207-83ee-45d8-9c98-5039a1a01207"
```

v6 JSON command output:

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

In the past versions, `spo eventreceiver get` returned an array with a single object. This has been adjusted, now the command will only return the object.

v5 JSON command output:

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

v6 JSON command output:

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

Update your scripts to expect a single object instead of an array.

## Updated `spo group user <verb>` commands

We've renamed the `spo group user <verb>` commands to `spo group member <verb>` to better cover all possible scenarios. In the near future we'll be adding support to add Azure AD Groups as group member. Using `spo group member` better fits the intended situation of adding either users or Azure AD groups.

As a side issue, we've also updated the response output of the `spo group member list` command in JSON output mode. This returned a member array within a parent `value` object. In the new situation, the command returns the array without the parent `value` object.

### What action do I need to take?

Update your scripts to use the new `member` noun instead of `user`. If you are using the output of `spo group member list` in JSON output mode, update your scripts and remove the `value` object.

## Removed short notation for option asAdmin in pp commands

We've decided to remove all short notations for option `--asAdmin` in pp commands. In previous versions, many commands had the notation `-a, --asAdmin`. This has been changed to `--adAdmin`, we removed the short notation to align it with our naming convention.

Affected commands:

- [pp card clone](./cmd/pp/card/card-clone.mdx)
- [pp card get](./cmd/pp/card/card-get.mdx)
- [pp card list](./cmd/pp/card/card-list.mdx)
- [pp card remove](./cmd/pp/card/card-remove.mdx)
- [pp dataverse table get](./cmd/pp/dataverse/dataverse-table-get.mdx)
- [pp dataverse table list](./cmd/pp/dataverse/dataverse-table-list.mdx)
- [pp dataverse table remove](./cmd/pp/dataverse/dataverse-table-remove.mdx)
- [pp environment get](./cmd/pp/environment/environment-get.mdx)
- [pp environment list](./cmd/pp/environment/environment-list.mdx)
- [pp solution get](./cmd/pp/solution/solution-get.mdx)
- [pp solution list](./cmd/pp/solution/solution-list.mdx)
- [pp solution remove](./cmd/pp/solution/solution-remove.mdx)
- [pp solution publisher get](./cmd/pp/solution/solution-publisher-get.mdx)
- [pp solution publisher list](./cmd/pp/solution/solution-publisher-list.mdx)
- [pp solution publisher remove](./cmd/pp/solution/solution-publisher-remove.mdx)

### What action do I need to take?

Update your scripts to use `--asAdmin` instead of `-a`.

## Updated `teams app list` command

The logic to list the installed apps in a specified team is moved to a new command [teams team app list](./cmd/teams/team/team-app-list.mdx). As a result, the command [teams app list](./cmd/teams/app/app-list.mdx) only displays the installed apps from the Microsoft Teams app catalog. The command [teams app list](./cmd/teams/app/app-list.mdx) does no longer contain the options `all`, `teamId` and `teamName`. In addition, there is a new option for this command that allows you to indicate which installed apps from the Microsoft Teams app catalog you want to list according to the distribution method.

### What action do I need to take?

Update your scripts to use the [teams app list](./cmd/teams/app/app-list.mdx) command if you want to list the installed apps in the Microsoft Teams app catalog. If you want to list the installed apps in a specified team, use the [teams team app list](./cmd/teams/team/team-app-list.mdx) command instead.

## Aligned options with naming convention

As we've been adding more commands to the CLI, we noticed that several commands were using inconsistent options names. Our naming convention states that options that refer to the last noun in the command, don't need that noun as a prefix. for example: the option `--webUrl` for `m365 spo web list` has been renamed to `--url` as the last noun is `web`. In version 6 of the CLI for Microsoft 365, we updated all these options to be consistent and make it easier for you to use the CLI.

We've updated the following commands and options:

Command|Old option|New option
--|--|--
[aad approleassignment add](./cmd/aad/approleassignment/approleassignment-add.mdx)|`objectId`|`appObjectId`
[aad approleassignment add](./cmd/aad/approleassignment/approleassignment-add.mdx)|`displayName`|`appDisplayName`
[aad approleassignment list](./cmd/aad/approleassignment/approleassignment-list.mdx)|`objectId`|`appObjectId`
[aad approleassignment list](./cmd/aad/approleassignment/approleassignment-list.mdx)|`displayName`|`appDisplayName`
[aad approleassignment remove](./cmd/aad/approleassignment/approleassignment-remove.mdx)|`objectId`|`appObjectId`
[aad approleassignment remove](./cmd/aad/approleassignment/approleassignment-remove.mdx)|`displayName`|`appDisplayName`
[aad o365group add](./cmd/aad/o365group/o365group-add.mdx)|`isPrivate [isPrivate]`|`isPrivate` (changed to flag)
[aad o365group recyclebinitem list](./cmd/aad/o365group/o365group-recyclebinitem-list.mdx)|`displayName`|`groupDisplayName`
[aad o365group recyclebinitem list](./cmd/aad/o365group/o365group-recyclebinitem-list.mdx)|`mailNickname`|`groupMailNickname`
[aad o365group teamify](./cmd/aad/o365group/o365group-teamify.mdx)|`groupId`|`id`
[aad policy list](./cmd/aad/policy/policy-list.mdx)|`policyType`|`type`
[aad sp get](./cmd/aad/sp/sp-get.mdx)|`displayName`|`appDisplayName`
[aad sp get](./cmd/aad/sp/sp-get.mdx)|`objectId`|`appObjectId`
[flow disable](./cmd/flow/flow-disable.mdx)|`environment`|`environmentName`
[flow enable](./cmd/flow/flow-enable.mdx)|`environment`|`environmentName`
[flow export](./cmd/flow/flow-export.mdx)|`environment`|`environmentName`
[flow get](./cmd/flow/flow-get.mdx)|`environment`|`environmentName`
[flow list](./cmd/flow/flow-list.mdx)|`environment`|`environmentName`
[flow remove](./cmd/flow/flow-remove.mdx)|`environment`|`environmentName`
[flow run cancel](./cmd/flow/run/run-cancel.mdx)|`flow`|`flowName`
[flow run cancel](./cmd/flow/run/run-cancel.mdx)|`environment`|`environmentName`
[flow run get](./cmd/flow/run/run-get.mdx)|`flow`|`flowName`
[flow run get](./cmd/flow/run/run-get.mdx)|`environment`|`environmentName`
[flow run list](./cmd/flow/run/run-list.mdx)|`flow`|`flowName`
[flow run list](./cmd/flow/run/run-list.mdx)|`environment`|`environmentName`
[flow run resubmit](./cmd/flow/run/run-resubmit.mdx)|`flow`|`flowName`
[flow run resubmit](./cmd/flow/run/run-resubmit.mdx)|`environment`|`environmentName`
[outlook message move](./cmd/outlook/message/message-move.mdx)|`messageId`|`id`
[pa connector export](./cmd/pa/connector/connector-export.mdx)|`environment`|`environmentName`
[pa connector list](./cmd/pa/connector/connector-list.mdx)|`environment`|`environmentName`
[pa solution reference add](./cmd/pa/solution/solution-reference-add.mdx)|`path`|`projectPath`
[spfx package generate](./cmd/spfx/package/package-generate.md)|`packageName`|`name`
[spo app add](./cmd/spo/app/app-add.mdx)|`scope`|`appCatalogScope`
[spo app deploy](./cmd/spo/app/app-deploy.mdx)|`scope`|`appCatalogScope`
[spo app get](./cmd/spo/app/app-get.mdx)|`scope`|`appCatalogScope`
[spo app install](./cmd/spo/app/app-install.mdx)|`scope`|`appCatalogScope`
[spo app list](./cmd/spo/app/app-list.mdx)|`scope`|`appCatalogScope`
[spo app remove](./cmd/spo/app/app-remove.mdx)|`scope`|`appCatalogScope`
[spo app retract](./cmd/spo/app/app-retract.mdx)|`scope`|`appCatalogScope`
[spo app uninstall](./cmd/spo/app/app-uninstall.mdx)|`scope`|`appCatalogScope`
[spo app upgrade](./cmd/spo/app/app-upgrade.mdx)|`scope`|`appCatalogScope`
[spo apppage set](./cmd/spo/apppage/apppage-set.mdx)|`pageName`|`name`
[spo cdn policy list](./cmd/spo/cdn/cdn-policy-list.mdx)|`type`|`cdnType`
[spo cdn policy set](./cmd/spo/cdn/cdn-policy-set.mdx)|`type`|`cdnType`
[spo contenttype field set](./cmd/spo/contenttype/contenttype-field-set.mdx)|`fieldId`|`id`
[spo customaction add](./cmd/spo/customaction/customaction-add.mdx)|`url`|`webUrl`
[spo customaction clear](./cmd/spo/customaction/customaction-clear.mdx)|`url`|`webUrl`
[spo customaction get](./cmd/spo/customaction/customaction-get.mdx)|`url`|`webUrl`
[spo customaction list](./cmd/spo/customaction/customaction-list.mdx)|`url`|`webUrl`
[spo customaction remove](./cmd/spo/customaction/customaction-remove.mdx)|`url`|`webUrl`
[spo customaction set](./cmd/spo/customaction/customaction-set.mdx)|`url`|`webUrl`
[spo feature disable](./cmd/spo/feature/feature-disable.mdx)|`featureId`|`id`
[spo feature disable](./cmd/spo/feature/feature-disable.mdx)|`url`|`webUrl`
[spo feature enable](./cmd/spo/feature/feature-enable.mdx)|`featureId`|`id`
[spo feature enable](./cmd/spo/feature/feature-enable.mdx)|`url`|`webUrl`
[spo feature list](./cmd/spo/feature/feature-list.mdx)|`url`|`webUrl`
[spo field get](./cmd/spo/field/field-get.mdx)|`fieldTitle`|`title`
[spo field remove](./cmd/spo/field/field-remove.mdx)|`fieldTitle`|`title`
[spo file checkin](./cmd/spo/file/file-checkin.mdx)|`fileUrl`|`url`
[spo file checkout](./cmd/spo/file/file-checkout.mdx)|`fileUrl`|`url`
[spo file sharinginfo get](./cmd/spo/file/file-sharinginfo-get.mdx)|`url`|`fileUrl`
[spo file sharinginfo get](./cmd/spo/file/file-sharinginfo-get.mdx)|`id`|`fileId`
[spo folder get](./cmd/spo/folder/folder-get.mdx)|`folderUrl`|`url`
[spo folder remove](./cmd/spo/folder/folder-remove.mdx)|`folderUrl`|`url`
[spo folder rename](./cmd/spo/folder/folder-rename.mdx)|`folderUrl`|`url`
[spo site hubsite connect](./cmd/spo/site/site-hubsite-connect.mdx)|`url`|`siteUrl`
[spo site hubsite disconnect](./cmd/spo/site/site-hubsite-disconnect.mdx)|`url`|`siteUrl`
[spo hubsite register](./cmd/spo/hubsite/hubsite-register.mdx)|`url`|`siteUrl`
[spo hubsite rights grant](./cmd/spo/hubsite/hubsite-rights-grant.mdx)|`url`|`hubSiteUrl`
[spo hubsite rights revoke](./cmd/spo/hubsite/hubsite-rights-revoke.mdx)|`url`|`hubSiteUrl`
[spo knowledgehub set](./cmd/spo/knowledgehub/knowledgehub-set.mdx)|`url`|`siteUrl`
[spo list contenttype add](./cmd/spo/list/list-contenttype-add.mdx)|`contentTypeId`|`id`
[spo list contenttype remove](./cmd/spo/list/list-contenttype-remove.mdx)|`contentTypeId`|`id`
[spo list view field add](./cmd/spo/list/list-view-field-add.mdx)|`fieldId`|`id`
[spo list view field add](./cmd/spo/list/list-view-field-add.mdx)|`fieldTitle`|`title`
[spo list view field add](./cmd/spo/list/list-view-field-add.mdx)|`fieldPosition`|`position`
[spo list view field remove](./cmd/spo/list/list-view-field-remove.mdx)|`fieldId`|`id`
[spo list view field remove](./cmd/spo/list/list-view-field-remove.mdx)|`fieldTitle`|`title`
[spo list view field set](./cmd/spo/list/list-view-field-set.mdx)|`fieldId`|`id`
[spo list view field set](./cmd/spo/list/list-view-field-set.mdx)|`fieldTitle`|`title`
[spo list view field set](./cmd/spo/list/list-view-field-set.mdx)|`fieldPosition`|`position`
[spo list view get](./cmd/spo/list/list-view-get.mdx)|`viewId`|`id`
[spo list view get](./cmd/spo/list/list-view-get.mdx)|`viewTitle`|`title`
[spo list view remove](./cmd/spo/list/list-view-remove.mdx)|`viewId`|`id`
[spo list view remove](./cmd/spo/list/list-view-remove.mdx)|`viewTitle`|`title`
[spo list view set](./cmd/spo/list/list-view-set.mdx)|`viewId`|`id`
[spo list view set](./cmd/spo/list/list-view-set.mdx)|`viewTitle`|`title`
[spo listitem list](./cmd/spo/listitem/listitem-list.mdx)|`id`|`listId`
[spo listitem list](./cmd/spo/listitem/listitem-list.mdx)|`title`|`listTitle`
[spo listitem record declare](./cmd/spo/listitem/listitem-record-declare.mdx)|`id`|`listItemId`
[spo listitem record undeclare](./cmd/spo/listitem/listitem-record-undeclare.mdx)|`id`|`listItemId`
[spo page column get](./cmd/spo/page/page-column-get.mdx)|`name`|`pageName`
[spo page column list](./cmd/spo/page/page-column-list.mdx)|`name`|`pageName`
[spo page control get](./cmd/spo/page/page-control-get.mdx)|`name`|`pageName`
[spo page control list](./cmd/spo/page/page-control-list.mdx)|`name`|`pageName`
[spo page control set](./cmd/spo/page/page-control-set.mdx)|`name`|`pageName`
[spo page section add](./cmd/spo/page/page-section-add.mdx)|`name`|`pageName`
[spo page section get](./cmd/spo/page/page-section-get.mdx)|`name`|`pageName`
[spo page section list](./cmd/spo/page/page-section-list.mdx)|`name`|`pageName`
[spo serviceprincipal grant revoke](./cmd/spo/serviceprincipal/serviceprincipal-grant-revoke.mdx)|`grantId`|`id`
[spo serviceprincipal permissionrequest approve](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-approve.mdx)|`requestId`|`id`
[spo serviceprincipal permissionrequest deny](./cmd/spo/serviceprincipal/serviceprincipal-permissionrequest-deny.mdx)|`requestId`|`id`
[spo site appcatalog add](./cmd/spo/site/site-appcatalog-add.mdx)|`url`|`siteUrl`
[spo site appcatalog remove](./cmd/spo/site/site-appcatalog-remove.mdx)|`url`|`siteUrl`
[spo site apppermission get](./cmd/spo/site/site-apppermission-get.mdx)|`permissionId`|`id`
[spo site apppermission remove](./cmd/spo/site/site-apppermission-remove.mdx)|`permissionId`|`id`
[spo site apppermission set](./cmd/spo/site/site-apppermission-set.mdx)|`permissionId`|`id`
[spo site chrome set](./cmd/spo/site/site-chrome-set.md)|`url`|`siteUrl`
[spo site groupify](./cmd/spo/site/site-groupify.mdx)|`siteUrl`|`url`
[spo site rename](./cmd/spo/site/site-rename.mdx)|`siteUrl`|`url`
[spo sitedesign rights grant](./cmd/spo/sitedesign/sitedesign-rights-grant.mdx)|`id`|`siteDesignId`
[spo sitedesign rights list](./cmd/spo/sitedesign/sitedesign-rights-list.mdx)|`id`|`siteDesignId`
[spo sitedesign rights revoke](./cmd/spo/sitedesign/sitedesign-rights-revoke.mdx)|`id`|`siteDesignId`
[spo sitedesign task get](./cmd/spo/sitedesign/sitedesign-task-get.mdx)|`taskId`|`id`
[spo sitedesign task remove](./cmd/spo/sitedesign/sitedesign-task-remove.mdx)|`taskId`|`id`
[spo tenant recyclebinitem remove](./cmd/spo/tenant/tenant-recyclebinitem-remove.mdx)|`url`|`siteUrl`
[spo tenant recyclebinitem restore](./cmd/spo/tenant/tenant-recyclebinitem-restore.mdx)|`url`|`siteUrl`
[spo web add](./cmd/spo/web/web-add.mdx)|`webUrl`|`url`
[spo web get](./cmd/spo/web/web-get.mdx)|`webUrl`|`url`
[spo web list](./cmd/spo/web/web-list.mdx)|`webUrl`|`url`
[spo web reindex](./cmd/spo/web/web-reindex.mdx)|`webUrl`|`url`
[spo web remove](./cmd/spo/web/web-remove.mdx)|`webUrl`|`url`
[spo web set](./cmd/spo/web/web-set.mdx)|`webUrl`|`url`
[teams app app install](./cmd/teams/app/app-install.mdx)|`appId`|`id`
[teams app app uninstall](./cmd/teams/app/app-uninstall.mdx)|`appId`|`id`
[teams channel get](./cmd/teams/channel/channel-get.mdx)|`channelId`|`id`
[teams channel get](./cmd/teams/channel/channel-get.mdx)|`channelName`|`name`
[teams channel remove](./cmd/teams/channel/channel-remove.mdx)|`channelId`|`id`
[teams channel remove](./cmd/teams/channel/channel-remove.mdx)|`channelName`|`name`
[teams channel set](./cmd/teams/channel/channel-set.mdx)|`channelName`|`name`
[teams message get](./cmd/teams/message/message-get.mdx)|`messageId`|`id`
[teams tab get](./cmd/teams/tab/tab-get.mdx)|`tabId`|`id`
[teams tab get](./cmd/teams/tab/tab-get.mdx)|`tabName`|`name`
[teams tab remove](./cmd/teams/tab/tab-remove.mdx)|`tabId`|`id`
[teams team archive](./cmd/teams/team/team-archive.mdx)|`teamId`|`id`
[teams team clone](./cmd/teams/team/team-clone.mdx)|`teamId`|`id`
[teams team remove](./cmd/teams/team/team-remove.mdx)|`teamId`|`id`
[teams team set](./cmd/teams/team/team-set.mdx)|`teamId`|`id`
[teams team unarchive](./cmd/teams/team/team-unarchive.mdx)|`teamId`|`id`
[teams user app add](./cmd/teams/user/user-app-add.mdx)|`appId`|`id`
[teams user app remove](./cmd/teams/user/user-app-remove.mdx)|`appId`|`id`
[viva connections app create](./cmd/viva/connections/connections-app-create.mdx)|`appName`|`name`
[yammer group user add](./cmd/yammer/group/group-user-add.mdx)|`id`|`groupId`
[yammer group user add](./cmd/yammer/group/group-user-add.mdx)|`userId`|`id`
[yammer group user remove](./cmd/yammer/group/group-user-remove.mdx)|`id`|`groupId`
[yammer group user remove](./cmd/yammer/group/group-user-remove.mdx)|`userId`|`id`
[yammer message like set](./cmd/yammer/message/message-like-set.mdx)|`id`|`messageId`
[yammer user get](./cmd/yammer/user/user-get.mdx)|`userId`|`id`

### What action do I need to take?

If you use any of the commands listed above, ensure that you use the new option names.
