# teams channel set

Updates properties of the specified channel in the given Microsoft Teams team

## Usage

```sh
m365 teams channel set [options]
```

## Options

`--teamId [teamId]`
: The ID of the team where the channel to update is located. Specify either `teamId` or `teamName` but not both.

`--teamName [teamName]`
: The display name of the team where the channel to update is located. Specify either `teamId` or `teamName` but not both.

`-i, --id [id]`
: The ID of the channel to update. Specify either `id` or `name` but not both.

`--name [name]`
: The name of the channel to update. Specify either `id` or `name` but not both.

`--newName [newName]`
: The new name of the channel.

`--description [description]`
: The description of the channel.

--8<-- "docs/cmd/_global.md"

## Examples
  
Set new description and display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --name Reviews --newName Projects --description "Channel for new projects"
```

Set new display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --name Reviews --newName Projects
```

## Response

The command won't return a response on success.
