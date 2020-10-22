# spo hidedefaultthemes set

Sets the value of the HideDefaultThemes setting

## Usage

```sh
m365 spo hidedefaultthemes set [options]
```

## Options

`--hideDefaultThemes <hideDefaultThemes>`
: Set to `true` to hide default themes and to `false` to show them

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Hide default themes and allow users to use organization themes only

```sh
m365 spo hidedefaultthemes set --hideDefaultThemes true
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)