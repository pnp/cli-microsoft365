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

`--type [type]`
: Filter the results to only channels of a given type: `standard, private`. By default all channels are listed.

--8<-- "docs/cmd/_global.md"

## Examples
  
List all channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000
```

List all channels in a specified Microsoft Teams team with name _Team Name_

```sh
m365 teams channel list --teamName "Team Name"
```

List private channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000 --type private
```