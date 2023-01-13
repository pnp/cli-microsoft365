# spo cdn policy list

Lists CDN policies settings for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn policy list [options]
```

## Options

`-t, --cdnType [cdnType]`
: Type of CDN to manage. `Public,Private`. Default `Public`

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --cdnType` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Show the list of policies configured for the Public CDN

```sh
m365 spo cdn policy list
```

Show the list of policies configured for the Private CDN

```sh
m365 spo cdn policy list --cdnType Private
```

## More information

- Use Microsoft 365 CDN with SharePoint Online: [https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide](https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide)
