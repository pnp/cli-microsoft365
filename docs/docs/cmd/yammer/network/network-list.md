# Yammer network list

Returns a list of networks to which the current user has access

## Usage

```sh
m365 yammer network list [options]
```

## Options

`--includeSuspended`
: Include the networks in which the user is suspended

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns the current user's networks

```sh
m365 yammer network list
```

Returns the current user's networks including the networks in which the user is suspended

```sh
m365 yammer network list --includeSuspended
```
