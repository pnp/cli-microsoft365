# yammer group user add

Adds a user from a Yammer group

## Usage

```sh
yammer group user add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--id <userId>`|The Group ID to process
`--userId [userId]`|Adds the user with the specified ID to the Yammer Group. Defaults to the current user
`--email [email]`|Adds the user with the specified e-mail to the Yammer Group. It will return a HTTP 400 error message if the user is not a member of the network
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the Office 365 CLI the permission to the Yammer API. To do this, execute the `consent --service yammer` command.

## Examples

Adds the current user to the group with the ID `5611239081`

```sh
yammer group user add --id 5611239081
```

Adds the the user with the ID `66622349` to the group with the ID `5611239081`

```sh
yammer group user add --id 5611239081 --userId 66622349
```

Adds the the user with the e-mail `suzy@contoso.com` to the group with the ID `5611239081`

```sh
yammer group user add --id 5611239081 --email suzy@contoso.com
``` 