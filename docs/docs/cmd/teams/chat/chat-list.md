# teams chat list

Lists all Microsoft Teams chat conversations for the current user.

## Usage

```sh
m365 teams chat list [options]
```

## Options

`-t, --type [chatType]`
: The chat type to optionally filter chat conversations by type. The value can be `oneOnOne`, `group` or `meeting`.

--8<-- "docs/cmd/_global.md"

## Examples

List all the Microsoft Teams chat conversations of the current user.

```sh
m365 teams chat list
```

List only the one on one Microsoft Teams chat conversations.

```sh
m365 teams chat list --type oneOnOne
```
