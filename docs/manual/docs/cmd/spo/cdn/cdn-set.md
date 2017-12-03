# spo cdn set

Enable or disable the specified Office 365 CDN

## Usage

```sh
spo cdn set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-e, --enabled <enabled>`|Set to true to enable CDN or to false to disable it. Valid values are true|false
`-t, --type [type]`|Type of CDN to manage. `Public|Private`. Default `Public`
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To enable or disable an Office 365 CDN, you have to first connect to a tenant admin site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.
If you are connected to a different site and will try to manage tenant properties,
you will get an error.

Using the `-t, --type` option you can choose whether you want to manage the settings of
the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

Using the `-e, --enabled` option you can specify whether the given CDN type should be
enabled or disabled. Use true to enable the specified CDN and false to
disable it.

## Examples

Enable the Office 365 Public CDN on the current tenant

```sh
spo cdn set -t Public -e true
```

Disable the Office 365 Public CDN on the current tenant

```sh
spo cdn set -t Public -e false
```

## More information

- General availability of Office 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
