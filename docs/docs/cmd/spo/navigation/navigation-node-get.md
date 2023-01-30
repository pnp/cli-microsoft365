# spo navigation node get

Gets information about a specific navigation node.

## Usage

```sh
m365 spo navigation node get [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site.

`--id <id>`
: Id of the navigation node.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve information for a specific navigation node.

```sh
m365 spo navigation node get --webUrl https://contoso.sharepoint.com/sites/team-a --id 2209
```

## Response

=== "JSON"

    ```json
    {
      "AudienceIds": null,
      "CurrentLCID": 1033,
      "Id": 2209,
      "IsDocLib": true,
      "IsExternal": false,
      "IsVisible": true,
      "ListTemplateType": 100,
      "Title": "Work Status",
      "Url": "/sites/team-a/Lists/Work Status/AllItems.aspx"
    }
    ```

=== "Text"

    ```text
    AudienceIds     : null
    CurrentLCID     : 1033
    Id              : 2209
    IsDocLib        : true
    IsExternal      : false
    IsVisible       : true
    ListTemplateType: 100
    Title           : Work Status
    Url             : /sites/team-a/Lists/Work Status/AllItems.aspx
    ```

=== "CSV"

    ```csv
    AudienceIds,CurrentLCID,Id,IsDocLib,IsExternal,IsVisible,ListTemplateType,Title,Url
    ,1033,2209,1,,1,100,Work Status,/sites/team-a/Lists/Work Status/AllItems.aspx
    ```

=== "Markdown"

    ```md
    # spo navigation node get --webUrl "https://contoso.sharepoint.com/sites/team-a" --id "2209"

    Date: 1/29/2023

    ## Work Status (2209)

    Property | Value
    ---------|-------
    AudienceIds | null
    CurrentLCID | 1033
    Id | 2209
    IsDocLib | true
    IsExternal | false
    IsVisible | true
    ListTemplateType | 100
    Title | Work Status
    Url | /sites/team-a/Lists/Work Status/AllItems.aspx
    ```
