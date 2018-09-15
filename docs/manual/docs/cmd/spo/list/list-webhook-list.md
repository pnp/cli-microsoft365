# spo list webhook list

Lists all webhooks for the specified list

## Usage

```sh
spo list webhook list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list to retrieve webhooks for is located
`-i, --id [id]`|ID of the list to retrieve all webhooks for. Specify either `id` or `title` but not both
`-t, --title [title]`|Title of the list to retrieve all webhooks for. Specify either `id` or `title` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to SharePoint, using the [spo login](../login.md) command.

## Remarks

To list all webhooks for a list, you have to first log in to SharePoint using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

List all webhooks for a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo list webhook list --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all webhooks for a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo list webhook list --webUrl https://contoso.sharepoint.com/sites/project-x --title Documents
```