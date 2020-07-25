# teams team set

Updates settings of a Microsoft Teams team

## Usage

```sh
m365 teams team set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team for which to update settings

`--displayName [displayName]`
: The display name for the Microsoft Teams team

`--description [description]`
: The description for the Microsoft Teams team

`--mailNickName [mailNickName]`
: The mail alias for the Microsoft Teams team

`--classification [classification]`
: The classification for the Microsoft Teams team

`--visibility [visibility]`
: The visibility of the Microsoft Teams team. Valid values `Private,Public`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Set Microsoft Teams team visibility as Private

```sh
m365 teams team set --teamId '00000000-0000-0000-0000-000000000000' --visibility Private
```

Set Microsoft Teams team classification as MBI

```sh
m365 teams team set --teamId '00000000-0000-0000-0000-000000000000' --classification MBI
```