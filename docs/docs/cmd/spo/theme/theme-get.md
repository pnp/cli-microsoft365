# spo theme get

Gets custom theme information

## Usage

```sh
m365 spo theme get [options]
```

## Options

`-n, --name <name>`
: The name of the theme to retrieve

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Get information about a theme

```sh
m365 spo theme get --name Contoso-Blue
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)
