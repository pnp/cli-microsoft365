# spo sitedesign rights list

Gets a list of principals that have access to a site design

## Usage

```sh
m365 spo sitedesign rights list [options]
```

## Options

`-i, --id <id>`
: The ID of the site design to get rights information from

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` doesn't refer to an existing site script, you will get a `File not found` error.

If no permissions are listed, it means that the particular site design is visible to everyone.

## Examples

Get information about rights granted for the site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo sitedesign rights list --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
