# aad user password validate

Check a user's password against the organization's password validation policy

## Usage

```sh
m365 aad user password validate [options]
```

## Options

`-p, --password <password>`
: The password to be validated.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Validate password _cli365P@ssW0rd_ against the organization's password validation policy

```sh
m365 aad user password validate --password "cli365P@ssW0rd"
```