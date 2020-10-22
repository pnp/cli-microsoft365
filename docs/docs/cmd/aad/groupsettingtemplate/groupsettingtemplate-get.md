# aad groupsettingtemplate get

Gets information about the specified Azure AD group settings template

## Usage

```sh
m365 aad groupsettingtemplate get [options]
```

## Options

`-i, --id [id]`
: The ID of the settings template to retrieve. Specify the `id` or `displayName` but not both

`-n, --displayName [displayName]`
: The display name of the settings template to retrieve. Specify the `id` or `displayName` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the group setting template with id _62375ab9-6b52-47ed-826b-58e47e0e304b_

```sh
m365 aad groupsettingtemplate get --id 62375ab9-6b52-47ed-826b-58e47e0e304b
```

Get information about the group setting template with display name _Group.Unified_

```sh
m365 aad groupsettingtemplate get --displayName Group.Unified
```