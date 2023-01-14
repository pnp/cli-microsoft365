# spo cdn origin remove

Removes CDN origin for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn origin remove [options]
```

## Options

`-t, --type [type]`
: Type of CDN to manage. `Public,Private`. Default `Public`

`-r, --origin <origin>`
: Origin to remove from the current CDN configuration

`--confirm`
: Don't prompt for confirming removal of a tenant property

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --type` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Remove _*/CDN_ from the list of origins of the Public CDN

```sh
m365 spo cdn origin remove --type Public --origin */CDN
```

## More information

- Use Microsoft 365 CDN with SharePoint Online: [https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide](https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide)
