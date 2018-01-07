# spo cdn policy set

Sets CDN policy value for the current SharePoint Online tenant

## Usage

```sh
spo cdn policy set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-t, --type [type]`|Type of CDN to manage. `Public|Private`. Default `Public`
`-p, --policy <policy>`|CDN policy to configure. `IncludeFileExtensions|ExcludeRestrictedSiteClassifications`
`-v, --value <value>`|Value for the policy to configure
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To set the policy of an Office 365 CDN, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to manage tenant properties, you will get an error.

Using the `-t, --type` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Set the list of extensions supported by the Public CDN

```sh
spo cdn policy set -t Public -p IncludeFileExtensions -v CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON
```

## More information

- General availability of Office 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)