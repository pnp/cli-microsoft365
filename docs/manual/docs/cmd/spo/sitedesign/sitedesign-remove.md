# spo sitedesign remove

Removes the specified site design

## Usage

```sh
spo sitedesign remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|Site design ID
`--confirm`|Don't prompt for confirming removing the site design
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To remove a site design, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `id` doesn't refer to an existing site design, you will get a `File not found` error.

## Examples

Remove site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_. Will prompt for confirmation before removing the design

```sh
spo sitedesign remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Remove site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ without prompting for confirmation

```sh
spo sitedesign remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --confirm
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)