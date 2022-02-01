# teams channel primary get

Gets information about the primary channel of a Microsoft Teams team

## Usage

```sh
m365 teams channel primary get [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team to retrieve the primary channel. Specify either teamId or teamName but not both

`--teamName [teamName]`
: The display name of the team to retrieve the primary channel. Specify either teamId or teamName but not both

--8<-- "docs/cmd/_global.md"

## Examples
  
Get information about primary channel of a Microsoft Teams team 

```sh
m365 teams channel primary get --teamName "Team Name" 
```