# spo sitedesign remove

Removes the specified site design

## Usage

```sh
m365 spo sitedesign remove [options]
```

## Options

`-i, --id <id>`
: Site design ID

`--confirm`
: Don't prompt for confirming removing the site design

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` doesn't refer to an existing site design, you will get a `File not found` error.

## Examples

Remove site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_. Will prompt for confirmation before removing the design

```sh
m365 spo sitedesign remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Remove site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ without prompting for confirmation

```sh
m365 spo sitedesign remove --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --confirm
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)