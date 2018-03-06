# spo sitedesign apply

Applies a site design to an existing site collection

## Usage

```sh
spo sitedesign apply [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the site design to apply
`-u, --webUrl <webUrl>`|The URL of the site to apply the site design to
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To apply a site design to an existing site collection, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Apply the site design with ID 9b142c22-037f-4a7f-9017-e9d8c0e34b98 to the site collection https://contoso.sharepoint.com/sites/project-x

```sh
spo sitedesign apply --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)