# spo orgassetslibrary list

List all libraries that are assigned as asset library

## Usage

```sh
m365 spo orgassetslibrary list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List all libraries that are assigned as asset library

```sh
m365 spo orgassetslibrary list
```

## Response

### Standard response

=== "JSON"

    ```json
    {
      "Url": "/",
      "Libraries": [
        {
          "DisplayName": "Site Assets",
          "LibraryUrl": "SiteAssets",
          "ListId": "/Guid(0a327c3f-ba82-4b19-bfa1-628405539420)/",
          "ThumbnailUrl": null
        }
      ]
    }
    ```

=== "Text"

    ```text
    Libraries: [{"DisplayName":"Site Assets","LibraryUrl":"SiteAssets","ListId":"/Guid(0a327c3f-ba82-4b19-bfa1-628405539420)/","ThumbnailUrl":null}]
    Url      : /
    ```

=== "CSV"

    ```csv
    Url,Libraries
    /,"[{""DisplayName"":""Site Assets"",""LibraryUrl"":""SiteAssets"",""ListId"":""/Guid(0a327c3f-ba82-4b19-bfa1-628405539420)/"",""ThumbnailUrl"":null}]"
    ```

### Response when no library is assigned as asset library

=== "Text"

    ```text
    No libraries in Organization Assets
    ```
