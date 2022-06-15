# aad o365group recyclebinitem remove

Permanently deletes a Microsoft 365 Group from the recycle bin in the current tenant

## Usage

```sh
m365 aad o365group recyclebinitem remove [options]
```

## Options

`-i, --id [id]`
: The ID of the Microsoft 365 Group to remove. Specify either `id`, `displayName` or `mailNickname` but not multiple.

`-d, --displayName [displayName]`
: Display name for the Microsoft 365 Group to remove. Specify either `id`, `displayName` or `mailNickname` but not multiple.

`-m, --mailNickname [mailNickname]`
: Name of the group e-mail (part before the @). Specify either `id`, `displayName` or `mailNickname` but not multiple.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes the Microsoft 365 Group with specific ID

```sh
m365 aad o365group recyclebinitem remove --id "00000000-0000-0000-0000-000000000000"
```

Removes the Microsoft 365 Group with specific name

```sh
m365 aad o365group recyclebinitem remove --displayName "My Group"
```

Remove the Microsoft 365 Group with specific mail nickname without confirmation

```sh
m365 aad o365group recyclebinitem remove --mailNickname "Mygroup" --confirm
```
