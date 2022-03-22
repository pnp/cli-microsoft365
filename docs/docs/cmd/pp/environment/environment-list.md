# pp environment list

Lists Microsoft Power Platform environments

## Usage

```sh
m365 pp environment list [options]
```

## Options

`-a, --asAdmin [teamId]`
Run the command as admin and retrieve all environments. Lists only environments you have explicitly are assigned permissions to by default.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.
    Register CLI for Microsoft 365 or Azure AD application as a management application for the Power Platform using 
    m365 pp managementapp add [options] 

## Examples

List Microsoft Power Platform environments

```sh
m365 pp environment list
```

List all Microsoft Power Platform environments

```sh
m365 pp environment list --asAdmin
```
