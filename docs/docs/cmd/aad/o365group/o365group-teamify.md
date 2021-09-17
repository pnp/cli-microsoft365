# aad o365group teamify

Creates a new Microsoft Teams team under existing Microsoft 365 group

## Usage

```sh
m365 aad o365group teamify [options]
```

## Options

`-i, --groupId [groupId]`
: The ID of the Microsoft 365 Group to connect to Microsoft Teams. Specify either groupId or mailNickname but not both

`--mailNickname [mailNickname]`
: The mail alias of the Microsoft 365 Group to connect to Microsoft Teams. Specify either groupId or mailNickname but not both

--8<-- "docs/cmd/_global.md"

## Examples

Creates a new Microsoft Teams team under existing Microsoft 365 group with id _e3f60f99-0bad-481f-9e9f-ff0f572fbd03_

```sh
m365 aad o365group teamify --groupId e3f60f99-0bad-481f-9e9f-ff0f572fbd03
```

Creates a new Microsoft Teams team under existing Microsoft 365 group with mailNickname _GroupName_

```sh
m365 aad o365group teamify --mailNickname GroupName
```