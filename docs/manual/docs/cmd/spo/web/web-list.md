# spo web list

Gets all webs within the specified site

## Usage

```sh
spo web list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the lists to retrieve are located
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get all webs, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Return all webs located in site _https://contoso.sharepoint.com/_

```sh
spo web list -u https://contoso.sharepoint.com
```

## More information

- Web REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn499819.aspx#bk_WebCollection](https://msdn.microsoft.com/en-us/library/office/dn499819.aspx#bk_WebCollection)