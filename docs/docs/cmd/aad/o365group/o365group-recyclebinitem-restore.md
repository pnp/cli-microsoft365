# aad o365group restore

Restores a deleted Microsoft 365 Group

## Usage

```sh
m365 aad o365group recyclebinitem restore [options]
```

## Alias

```sh
m365 aad o365group restore [options]
```

## Options

`-i, --id [id]`
: The ID of the Microsoft 365 Group to restore. Specify either `id`, `displayName` or `mailNickname` but not multiple.

`-d, --displayName [displayName]`
: Display name for the Microsoft 365 Group to restore. Specify either `id`, `displayName` or `mailNickname` but not multiple.

`-m, --mailNickname [mailNickname]`
: Name of the group e-mail (part before the @). Specify either `id`, `displayName` or `mailNickname` but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Restores the Microsoft 365 Group with specific ID

```sh
m365 aad o365group recyclebinitem restore --id 28beab62-7540-4db1-a23f-29a6018a3848
```

Restores the Microsoft 365 Group with specific name

```sh
m365 aad o365group recyclebinitem restore --displayName "My Group"
```

Restores the Microsoft 365 Group with specific mail nickname

```sh
m365 aad o365group recyclebinitem restore --mailNickname "Mygroup"
```
