# spo app teamspackage download

Downloads Teams app package for an SPFx solution deployed to tenant app catalog

## Usage

```sh
m365 spo app teamspackage download [options]
```

## Options

`--appItemUniqueId [appItemUniqueId]`
: The unique ID of the SPFx app to download the Teams package for. Specify `appItemUniqueId`, `appItemId` or `appName`

`--appItemId [appItemId]`
: The ID of the list item behind the SPFx app to download the Teams package for. Specify `appItemUniqueId`, `appItemId` or `appName`

`--appName [appName]`
: The name of the sppkg file to download the Teams package for. Specify `appItemUniqueId`, `appItemId` or `appName`

`--fileName [fileName]`
: Name of the file to save the package to. If not specified will use the name of the sppkg file with a `.zip` extension

`-u, --appCatalogUrl [appCatalogUrl]`
: URL of the tenant app catalog. If not specified, the command will try to autodiscover it

--8<-- "docs/cmd/_global.md"

## Remarks

Download the Teams app package for an SPFx solution works only for solutions deployed to the tenant app catalog.

If you try to download Teams app package for an SPFx solution that doesn't support deployment to Teams, you'll get the _Request failed with status code 404_ error.

For maximum performance, specify the URL of the tenant app catalog, the item ID (`appItemId`) of the SPFx package for which you want to download the Teams app package and the name of the file where you want to save the downloaded package to (`fileName`).

## Examples

Downloads the Teams app package for the SPFx solution deployed to the tenant app catalog with the ID `1` to a file with .zip extension named after the .sppkg file:

```sh
m365 spo app teamspackage download --appItemId 1
```

Downloads the Teams app package for the SPFx solution deployed to the tenant app catalog with the unique item ID `335a5612-3e85-462d-9d5b-c014b5abeac5` to the specified file:

```sh
m365 spo app teamspackage download --appItemUniqueId 335a5612-3e85-462d-9d5b-c014b5abeac5 --fileName my-app.zip
```

Downloads the Teams app package for the SPFx solution deployed to the specified tenant app catalog:

```sh
m365 spo app teamspackage download --appName my-app.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/appcatalog
```
