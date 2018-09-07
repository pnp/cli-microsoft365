# spo app get

Gets information about the specific app from the tenant app catalog

## Usage

```sh
spo app get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id [id]`|ID of the app to retrieve information for. Specify the `id` or the `name` but not both
`-n, --name [name]`|Name of the app to retrieve information for. Specify the `id` or the `name` but not both
`-u, --appCatalogUrl [appCatalogUrl]`|URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get information about the specified app available in the tenant app catalog, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Return details about the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ available in the tenant app catalog.

```sh
spo app get -i b2307a39-e878-458b-bc90-03bc578531d6
```

Return details about the app with name _solution.sppkg_ available in the tenant app catalog. Will try to detect the app catalog URL

```sh
spo app get --name solution.sppkg
```

Return details about the app with name _solution.sppkg_ available in the tenant app catalog using the specified app catalog URL

```sh
spo app get --name solution.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/apps
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)