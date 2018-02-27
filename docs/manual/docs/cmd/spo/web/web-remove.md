# spo web remove

Delete specified subsite

## Usage

```sh
spo web remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the subsite to remove
`--confirm`|Do not prompt for confirmation before deleting the subsite
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To delete a subsite, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Delete subsite without prompting for confirmation

```sh
spo web remove --webUrl https://contoso.sharepoint.com/subsite --confirm
```