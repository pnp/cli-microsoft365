# teams channel list

Lists channels in the specified Microsoft Teams team

## Usage

```sh
m365 teams channel list [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team to list the channels of. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team to list the channels of. Specify either `teamId` or `teamName` but not both

--8<-- "docs/cmd/_global.md"

## Examples
  
List the channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000
```

List the channels in a specified Microsoft Teams team with name _Team Name_

```sh
m365 teams channel list --teamName "Team Name"
```