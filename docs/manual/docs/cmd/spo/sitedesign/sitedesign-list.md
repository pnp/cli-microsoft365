# spo sitedesign list

Lists available site designs for creating modern sites

## Usage

```sh
spo sitedesign list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To list available site designs, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

List available site designs

```sh
spo sitedesign list
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)