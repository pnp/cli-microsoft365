# teams team archive

Archives specified Microsoft Teams team

## Usage

```sh
m365 teams team archive [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team to archive

`--shouldSetSpoSiteReadOnlyForMembers`
: Sets the permissions for team members to read-only on the SharePoint Online site associated with the team

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Using this command, global admins and Microsoft Teams service admins can access teams that they are not a member of.

When a team is archived, users can no longer send or like messages on any channel in the team, edit the team's name, description, or other settings, or in general make most changes to the team. Membership changes to the team continue to be allowed.

## Examples

Archive the specified Microsoft Teams team

```sh
m365 teams team archive --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
```

Archive the specified Microsoft Teams team and set permissions for team members to read-only on the SharePoint Online site associated with the team

```sh
m365 teams team archive --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --shouldSetSpoSiteReadOnlyForMembers
```