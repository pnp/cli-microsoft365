# yammer group list

Returns the list of groups in a Yammer network or the groups for a specific user

## Usage

```sh
m365 yammer group list [options]
```

## Options

`--userId [userId]`
: Returns the groups for a specific user

`--limit [limit]`
: Limits the groups returned

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns all Yammer network groups

```sh
m365 yammer group list
```

Returns all Yammer network groups for the user with the ID `5611239081`

```sh
m365 yammer group list --userId 5611239081
```

Returns the first 10 Yammer network groups

```sh
m365 yammer group list --limit 10
```

Returns the first 10 Yammer network groups for the user with the ID `5611239081`

```sh
m365 yammer group list --userId 5611239081 --limit 10
```
