# spo sitedesign get

Gets information about the specified site design

## Usage

```sh
m365 spo sitedesign get [options]
```

## Options

`-i, --id [id]`
: Site design ID. Specify either id or title but not both

`--title [title]`
: Site design title. Specify either id or title but not both

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` or `title` doesn't refer to an existing site design, you will get a `File not found` error.

## Examples

Get information about the site design with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo sitedesign get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Get information about the site design with title _Contoso Site Design_

```sh
m365 spo sitedesign get --title "Contoso Site Design"
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
