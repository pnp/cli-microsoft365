# yammer group user remove

Removes a user from a Yammer group

## Usage

```sh
m365 yammer group user remove [options]
```

## Options

`-h, --help`
: output usage information

`--id <id>`
: The ID of the Yammer group

`--userId [userId]`
: ID of the user to remove from the group. If not specified, removes the current user

`--confirm`
: Don't prompt for confirmation before removing the user from the group

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Remove the current user from the group with the ID `5611239081`

```sh
m365 yammer group user remove --id 5611239081
```

Remove the user with the ID `66622349` from the group with the ID `5611239081`

```sh
m365 yammer group user remove --id 5611239081 --userId 66622349
```

Remove the user with the ID `66622349` from the group with the ID `5611239081` without asking for confirmation

```sh
m365 yammer group user remove --id 5611239081 --userId 66622349 --confirm
```
