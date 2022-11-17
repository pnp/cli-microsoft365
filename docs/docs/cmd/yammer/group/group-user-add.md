# yammer group user add

Adds a user to a Yammer Group

## Usage

```sh
m365 yammer group user add [options]
```

## Options

`--groupId <groupId>`
: The ID of the group to add the user to

`--id [id]`
: ID of the user to add to the group. If not specified, adds the current user

`--email [email]`
: E-mail of the user to add to the group

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

If the specified user is not a member of the network, the command will return an HTTP 400 error message.

## Examples

Adds the current user to the group with the ID `5611239081`

```sh
m365 yammer group user add --groupId 5611239081
```

Adds the user with ID `66622349` to the group with the ID `5611239081`

```sh
m365 yammer group user add --groupId 5611239081 --id 66622349
```

Adds the user with e-mail `suzy@contoso.com` to the group with ID `5611239081`

```sh
m365 yammer group user add --groupId 5611239081 --email suzy@contoso.com
```

## Response

The command won't return a response on success.
