# spo list get

Gets information about the specific list

## Usage

```sh
spo list get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list to retrieve is located
`-i, --id <id>`|ID of the list to retrieve information for. Specify either ID or TITLE but not both
`-t, --title <title>`|Title of the list to retrieve information for. Specify either ID or TITLE but not both
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To retrieve list, you have to first connect to a SharePoint Online site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Return information about a list with ID _0CD891EF-AFCE-4E55-B836-FCE03286CCCF_ located in site or site collection _https://contoso.sharepoint.com/sites/project-x_

```sh
spo list get -i 0CD891EF-AFCE-4E55-B836-FCE03286CCCF -u https://contoso.sharepoint.com/sites/project-x
```

Return information about a list with TITLE _Documents_ located in site or site collection _https://contoso.sharepoint.com/sites/project-x_

```sh
spo list get --t Documents --u https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)