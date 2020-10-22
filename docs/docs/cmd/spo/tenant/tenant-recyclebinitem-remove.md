# spo tenant recyclebinitem remove

Removes the specified deleted Site Collection from Tenant Recycle Bin

## Usage

```sh
m365 spo tenant recyclebinitem remove [options]
```

## Options

`-u, --url`
: URL of the site to remove

`--wait`
: Wait for the site collection to be removed before completing the command

`--confirm`
: Don't prompt for confirming removing the deleted site collection

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Removing a site collection is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully removed, you should use the `--wait` flag. When using this flag, the `m365 spo tenant recyclebinitem remove` command will keep running until it received confirmation from Microsoft 365 that the site has been fully removed.

## Examples

Removes the specified deleted site collection from tenant recycle bin

```sh
m365 spo tenant recyclebinitem remove --url https://contoso.sharepoint.com/sites/team
```

Removes the specified deleted site collection from tenant recycle bin and wait for the removing process to complete

```sh
m365 spo tenant recyclebinitem remove --url https://contoso.sharepoint.com/sites/team --wait
```
