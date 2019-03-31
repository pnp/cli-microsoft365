# graph teams archive

Archive the specified team. When a team is archived, users can no longer send or like messages on any channel in the team, edit the team's name, description, or other settings, or in general make most changes to the team. Membership changes to the team continue to be allowed.

## Usage

```sh
graph teams archive [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId [teamId]`|The ID of the team for which to list installed apps
`--shouldSetSpoSiteReadOnlyForMembers [shouldSetSpoSiteReadOnlyForMembers]`|Set to `true` to set permissions for team members to read-only on the Sharepoint Online site associated with the team
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To archive a Microsoft Team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

To archive a Microsoft Team, specify that team's ID using the `teamId` option.

Set `shouldSetSpoSiteReadOnlyForMembers` option to `true` to set permissions for team members to read-only on the Sharepoint Online site associated with the team.

You can only archive a Team as a global administrator.

## Examples

Archive the specified Microsoft Teams team

```sh
graph teams archive --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
```

Archive the specified Microsoft Teams team and set permissions for team members to read-only on the Sharepoint Online site associated with the team

```sh
graph teams archive --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --shouldSetSpoSiteReadOnlyForMembers true
```