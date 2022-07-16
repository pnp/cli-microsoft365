# spo group add

Creates group in the specified site

## Usage

```sh
m365 spo group add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the group should be added.

`-n, --name <name>`
: Name of the group to add.

`--description [description]`
: The description for the group.

`--allowMembersEditMembership [allowMembersEditMembership]`
: Who can edit the membership of the group? When `true` members can edit membership, otherwise only owners can do this. Default `false`.

`--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]`
: Who can view the membership of the group? When `false` everyone can view the group membership, otherwise only group members can. Default `true`.

`--allowRequestToJoinLeave [allowRequestToJoinLeave]`
: Specify whether to allow users to request membership in this group and allow users to request to leave the group. Default `false`.

`--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]`
: If auto-accept is enabled, users will automatically be added or removed when they make a request. Default `false`.

`--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]`
: All membership requests will be sent to the email address specified.

--8<-- "docs/cmd/_global.md"

## Examples

Create group with title and description

```sh
m365 spo group add --webUrl https://contoso.sharepoint.com/sites/project-x --name "Project leaders" --description "This group contains all project leaders"
```

Create group with membership requests

```sh
m365 spo group add --webUrl https://contoso.sharepoint.com/sites/project-x --name "Project leaders" --allowRequestToJoinLeave true --requestToJoinLeaveEmailSetting john.doe@contoso.com
```
