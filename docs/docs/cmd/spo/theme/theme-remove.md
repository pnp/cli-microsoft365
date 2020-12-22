# spo theme remove

Removes existing theme

## Usage

```sh
m365 spo theme remove [options]
```

## Options

`-n, --name <name>`
: Name of the theme to remove

`--confirm`
: Do not prompt for confirmation before removing theme

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Remove theme. Will prompt for confirmation before removing the theme

```sh
m365 spo theme remove --name Contoso-Blue
```

Remove theme without prompting for confirmation

```sh
m365 spo theme remove --name Contoso-Blue --confirm
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)