# teams team set

Updates settings of a Microsoft Teams team

## Usage

```sh
m365 teams team set [options]
```

## Options

`-i, --id [id]`
: The ID of the Microsoft Teams team for which to update settings

`--teamId [teamId]`
: (deprecated. Use `id` instead) The ID of the Microsoft Teams team for which to update settings

`-n, --name [name]`
: The display name for the Microsoft Teams team for which to update settings

`--displayName [displayName]`
: (deprecated. Use `name` instead) The display name for the Microsoft Teams team for which to update settings

`--description [description]`
: The description for the Microsoft Teams team

`--mailNickName [mailNickName]`
: The mail alias for the Microsoft Teams team

`--classification [classification]`
: The classification for the Microsoft Teams team

`--visibility [visibility]`
: The visibility of the Microsoft Teams team. Valid values `Private,Public`

--8<-- "docs/cmd/_global.md"

## Examples

Set Microsoft Teams team visibility as Private

```sh
m365 teams team set --id "00000000-0000-0000-0000-000000000000" --visibility Private
```

Set Microsoft Teams team classification as MBI

```sh
m365 teams team set --id "00000000-0000-0000-0000-000000000000" --classification MBI
```
