# spo list list

Gets all lists within the specified site

## Usage

```sh
spo list list [options]
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
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get all lists, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Return all lists located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo list list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)