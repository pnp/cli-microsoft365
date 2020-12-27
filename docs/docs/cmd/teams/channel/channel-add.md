# teams channel add

Adds a channel to the specified Microsoft Teams team

## Usage

```sh
m365 teams channel add [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team to add the channel to. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team to add the channel to. Specify either `teamId` or `teamName` but not both

`-n, --name <name>`
: The name of the channel to add

`-d, --description [description]`
: The description of the channel to add

--8<-- "docs/cmd/_global.md"

## Remarks

You can only add a channel to the Microsoft Teams team you are a member of.

## Examples

Add channel to the specified Microsoft Teams team with id 6703ac8a-c49b-4fd4-8223-28f0ac3a6402

```sh
m365 teams channel add --teamId 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 --name climicrosoft365 --description development
```

Add channel to the specified Microsoft Teams team with name _Team Name_

```sh
m365 teams channel add --teamName "Team Name" --name climicrosoft365 --description development
```
