# spo app list

Lists apps from the tenant app catalog

## Usage

```sh
spo app list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To list the apps available in the tenant app catalog, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

List all apps available in the tenant app catalog

```sh
spo app list
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)