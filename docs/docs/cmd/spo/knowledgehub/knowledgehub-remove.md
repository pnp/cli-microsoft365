# spo knowledgehub remove

Removes the Knowledge Hub Site setting for your tenant

## Usage

```sh
m365 spo knowledgehub remove [options]
```

## Options

`--confirm`
: Do not prompt for confirmation before removing the Knowledge Hub Site setting for your tenant

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes the Knowledge Hub Site setting for your tenant

```sh
m365 spo knowledgehub remove
```

Removes the Knowledge Hub Site setting for your tenant without confirmation

```sh
m365 spo knowledgehub remove --confirm
```
