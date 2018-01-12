# spo sitedesign get

Gets information about the specified site design

## Usage

```sh
spo sitedesign get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|Site design ID
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a site design, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `id` doesn't refer to an existing site design, you will get a `File not found` error.

## Examples

Get information about the site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
spo sitedesign get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)