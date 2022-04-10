# spo theme apply

Applies theme to the specified site

## Usage

```sh
m365 spo theme apply [options]
```

## Options

`-n, --name <name>`
: Name of the theme to apply

`-u, --webUrl <webUrl>`
: URL of the site to which the theme should be applied

`--sharePointTheme`
: Set to specify if the supplied theme name is a standard SharePoint theme

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Following standard SharePoint themes are supported by the CLI for Microsoft 365: Blue, Orange, Red, Purple, Green, Gray, Dark Yellow and Dark Blue.

## Examples

Apply theme to the specified site

```sh
m365 spo theme apply --name Contoso-Blue --webUrl https://contoso.sharepoint.com/sites/project-x
```

Apply a standard SharePoint theme to the specified site

```sh
m365 spo theme apply --name Blue --webUrl https://contoso.sharepoint.com/sites/project-x --sharePointTheme
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)
- Theme Generator: [https://aka.ms/themedesigner](https://aka.ms/themedesigner)
