# spo site classic remove

Removes the specified site

## Usage

```sh
spo site classic remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site to remove
`--skipRecycleBin`|Set to directly remove the site without moving it to the Recycle Bin
`--fromRecycleBin`|Set to remove the site from the Recycle Bin
`--wait`|Wait for the site to be removed before completing the command
`--confirm`|Don't prompt for confirming removing the site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To remove a classic site, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

Deleting and creating classic site collections is by default asynchronous and depending on the current state of Office 365, might take up to few minutes. If you're building a script with steps that require the site to be fully deleted, you should use the `--wait` flag. When using this flag, the `spo site classic remove` command will keep running until it received confirmation from Office 365 that the site has been fully deleted.

## Examples

Remove the specified site and place it in the Recycle Bin

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite
```

Remove the site without moving it to the Recycle Bin

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite --skipRecycleBin
```

Remove the previously deleted site from the Recycle Bin

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite --fromRecycleBin
```

Remove the site without moving it to the Recycle Bin and wait for completion

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite --wait --skipRecycleBin
```