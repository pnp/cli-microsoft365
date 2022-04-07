# spo cdn origin list

List CDN origins settings for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn origin list [options]
```

## Options

`-t, --type [type]`
: Type of CDN to manage. `Public,Private`. Default `Public`

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --type` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Show the list of origins configured for the Public CDN

```sh
m365 spo cdn origin list
```

Show the list of origins configured for the Private CDN

```sh
m365 spo cdn origin list --type Private
```

## More information

- General availability of Microsoft 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
