# spo sitedesign rights list

Gets a list of principals that have access to a site design

## Usage

```sh
spo sitedesign rights list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the site design to get rights information from
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about site design rights, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `id` doesn't refer to an existing site script, you will get a `File not found` error.

If no permissions are listed, it means that the particular site design is visible to everyone.

## Examples

Get information about rights granted for the site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
spo sitedesign rights list --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)