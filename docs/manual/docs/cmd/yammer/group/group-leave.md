# yammer group leave

Leave a Yammer group

## Usage

```sh
yammer group leave [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--id <userId>`|The Group ID of the group to leave
`--userId [userId]`|Remove the user with the ID specified. Defaults to the current user
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the Office 365 CLI the permission to the Yammer API. To do this, execute the `consent --service yammer` command.

## Examples

Remove the current user from the group with the ID `5611239081`

```sh
yammer group leave --id 5611239081
```

Remove the the user with the ID `66622349'` from the group with the ID `5611239081`

```sh
yammer group leave --id 5611239081 --userId 66622349
```
