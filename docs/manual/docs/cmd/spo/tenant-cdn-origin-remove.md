# spo tenant cdn origin remove

Removes CDN origin for the current SharePoint Online tenant

## Usage

```sh
spo tenant cdn origin remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --type [type]`|Type of CDN to manage. `Public|Private`. Default `Public`
`-o, --origin <origin>`|Origin to remove from the current CDN configuration
`--confirm`|Don't prompt for confirming removal of a tenant property
`--verbose`|Runs command with verbose logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](connect.md) command.

## Remarks

To remove an origin from an Office 365 CDN, you have to first connect to a tenant admin site using the
[spo connect](connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.
If you are connected to a different site and will try to manage tenant properties,
you will get an error.

Using the `-t, --type` option you can choose whether you want to manage the settings of
the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

```sh
spo tenant cdn origin remove -t Public -o */CDN
```

removes */CDN from the list of origins of the Public CDN

## More information

- General availability of Office 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
