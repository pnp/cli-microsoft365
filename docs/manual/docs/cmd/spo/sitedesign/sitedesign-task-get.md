# spo sitedesign task get

Gets information about the specified site design scheduled for execution

## Usage

```sh
spo sitedesign task get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --taskId <taskId>`|The ID of the site design task to get information for
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get information about the specified site design scheduled for execution, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Get information about the specified site design scheduled for execution

```sh
spo sitedesign task get --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)