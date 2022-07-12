# teams team archive

Archives specified Microsoft Teams team

## Usage

```sh
m365 teams team archive [options]
```

## Options

`-i, --id [id]`
: The ID of the Microsoft Teams team to archive. Specify either id or name but not both

`-n, --name [name]`
: The display name of the Microsoft Teams team to archive. Specify either id or name but not both

`--teamId [teamId]`
: (deprecated. Use `id` instead) The ID of the Microsoft Teams team to archive

`--shouldSetSpoSiteReadOnlyForMembers`
: Sets the permissions for team members to read-only on the SharePoint Online site associated with the team

--8<-- "docs/cmd/_global.md"

## Remarks

Using this command, global admins and Microsoft Teams service admins can access teams that they are not a member of.

If the command finds multiple Microsoft Teams teams with the specified name, it will prompt you to disambiguate which team it should use, listing the discovered team IDs.

When a team is archived, users can no longer send or like messages on any channel in the team, edit the team's name, description, or other settings, or in general make most changes to the team. Membership changes to the team continue to be allowed.


## Examples

Archive the specified Microsoft Teams team with id _6f6fd3f7-9ba5-4488-bbe6-a789004d0d55_

```sh
m365 teams team archive --id 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
```

Archive the specified Microsoft Teams team with name _Team Name_

```sh
m365 teams team archive --name "Team Name"
```

Archive the specified Microsoft Teams team and set permissions for team members to read-only on the SharePoint Online site associated with the team

```sh
m365 teams team archive --id 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --shouldSetSpoSiteReadOnlyForMembers
```
