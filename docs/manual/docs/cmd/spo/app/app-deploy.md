# spo app deploy

Deploys the specified app in the tenant app catalog

## Usage

```sh
spo app deploy [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to deploy. Needs to be available in the tenant app catalog.
`-u, --appCatalogUrl [appCatalogUrl]`|(optional) URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To deploy an app in the tenant app catalog, you have to first connect to a SharePoint site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If you don't specify the URL of the tenant app catalog site using the **appCatalogUrl** option, the CLI will try to determine its URL automatically. This will be done using SharePoint Search. If the tenant app catalog site hasn't been crawled yet, the CLI will not find it and will prompt you to provide the URL yourself.

If the app with the specified ID doesn't exist in the tenant app catalog, the command will fail with an error. Before you can deploy an app, you have to add it to the tenant app catalog first using the [spo app add](./app-add.md) command.

## Examples

```sh
spo app deploy -i 058140e3-0e37-44fc-a1d3-79c487d371a3
```

Deploys the specified app in the tenant app catalog. Will try to resolve the URL of the tenant app catalog automatically.

```sh
spo app deploy -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/apps
```

Deploys the specified app in the tenant app catalog located at _https://contoso.sharepoint.com/sites/apps_

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)