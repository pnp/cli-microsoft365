# spo page remove

Removes a modern page

## Usage

```sh
spo page remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the page to remove
`-u, --webUrl <webUrl>`|URL of the site from which the page should be removed
`--confirm`|Do not prompt for confirmation before removing the page
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To remove new modern page, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If you try to remove a page with that does not exist, you
will get a `The file does not exist` error.

If you set the `--confirm` flag, you will not be prompted for confirmation before the page is actually removed.


## Examples

Remove a modern page.
```sh
spo page remove --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team
```

Remove a modern page without a confirmation prompt.
```sh
spo page remove --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --confirm
```