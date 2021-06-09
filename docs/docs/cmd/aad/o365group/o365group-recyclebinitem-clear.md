# aad o365group recyclebinitem clear

Clears Microsoft 365 Groups from the recycle bin in the current tenant

## Usage

```sh
m365 aad o365group recyclebinitem clear [options]
```

## Options

`--confirm`
: Don't prompt for confirmation to clear the recycle bin items.

--8<-- "docs/cmd/_global.md"

## Examples

Removes all deleted Microsoft 365 Groups in the tenant. Will prompt for confirmation before permanently removing groups.

```sh
m365 aad o365group recyclebinitem clear
```

Removes all deleted Microsoft 365 Groups in the tenant. Will NOT prompt for confirmation before permanently removing groups.

```sh
m365 aad o365group recyclebinitem clear --confirm
```