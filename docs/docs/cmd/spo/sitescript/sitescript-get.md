# spo sitescript get

Gets information about the specified site script

## Usage

```sh
m365 spo sitescript get [options]
```

## Options

`-i, --id <id>`
: Site script ID

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `id` doesn't refer to an existing site script, you will get a `File not found` error.

## Examples

Get information about the site script with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo sitescript get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint site design and site script overview: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview)
