# spo group set

Updates a group in the specified site

## Usage

```sh
m365 spo group set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the group is located.

`-i, --id [id]`
: ID of the group to update. Specify either `id` or `name` but not both.

`-n, --name [name]`
: Name of the group. Specify either `id` or `name` but not both.

`--newName [newName]`
: New name for the group.

`--description [description]`
: The description for the group.

`--allowMembersEditMembership [allowMembersEditMembership]`
: Who can edit the membership of the group? When `true` members can edit membership, otherwise only owners can do this.

`--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]`
: Who can view the membership of the group? When `true` only group members can view the membership, otherwise everyone can.

`--allowRequestToJoinLeave [allowRequestToJoinLeave]`
: Specify whether to allow users to request membership in this group and allow users to request to leave the group.

`--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]`
: If auto-accept is enabled, users will automatically be added or removed when they make a request.

`--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]`
: All membership requests will be sent to the email address specified.

`--ownerEmail [ownerEmail]`
: Set user with this email as owner of the group. Specify either `ownerEmail` or `ownerUserName` but not both.

`--ownerUserName [ownerUserName]`
: Set user with this login name as owner of the group. Specify either `ownerEmail` or `ownerUserName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Update group title and description

```sh
m365 spo group set --webUrl https://contoso.sharepoint.com/sites/project-x --id 18 --newTitle "Project leaders" --description "This group contains all project leaders"
```

Update group with membership requests

```sh
m365 spo group set --webUrl https://contoso.sharepoint.com/sites/project-x --title "Project leaders" --allowRequestToJoinLeave true --requestToJoinLeaveEmailSetting john.doe@contoso.com
```

Sets a specified user as group owner

```sh
m365 spo group set --webUrl https://contoso.sharepoint.com/sites/project-x --id 18 --ownerEmail john.doe@contoso.com
```
