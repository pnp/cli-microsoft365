# teams channel list

Lists channels in the specified Microsoft Teams team

## Usage

```sh
m365 teams channel list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId [teamId]`
: The ID of the team to list the channels of. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team to list the channels of. Specify either `teamId` or `teamName` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples
  
List the channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000
```

List the channels in a specified Microsoft Teams team with name _Team Name_

```sh
m365 teams channel list --teamName "Team Name"
```