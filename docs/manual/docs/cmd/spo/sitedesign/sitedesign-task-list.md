# spo sitedesign task list

Lists site designs scheduled for execution on the specified site

## Usage

```sh
spo sitedesign task list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site for which to list site designs scheduled for execution
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To list information about site designs scheduled for execution on the specified site, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

List site designs scheduled for execution on the specified site

```sh
spo sitedesign task list --webUrl https://contoso.sharepoint.com/sites/team-a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)