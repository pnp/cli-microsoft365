# spo cdn get

View current status of the specified Microsoft 365 CDN

## Usage

```sh
m365 spo cdn get [options]
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

Show if the Public CDN is currently enabled or not

```sh
m365 spo cdn get
```

Show if the Private CDN is currently enabled or not

```sh
m365 spo cdn get --type Private
```

## More information

- General availability of Microsoft 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
