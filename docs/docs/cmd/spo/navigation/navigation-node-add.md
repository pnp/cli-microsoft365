# spo navigation node add

Adds a navigation node to the specified site navigation

## Usage

```sh
m365 spo navigation node add [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site to which navigation should be modified

`-l, --location [location]`
: Navigation type where the node should be added. Available options: `QuickLaunch`, `TopNavigationBar`

`-t, --title <title>`
: Navigation node title

`--url <url>`
: Navigation node URL

`--parentNodeId [parentNodeId]`
: ID of the node below which the node should be added

`--isExternal`
: Set, if the navigation node points to an external URL

--8<-- "docs/cmd/_global.md"

## Examples

Add a navigation node pointing to a SharePoint page to the top navigation

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar --title About --url /sites/team-s/sitepages/about.aspx
```

Add a navigation node pointing to an external page to the quick launch

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch --title "About us" --url https://contoso.com/about-us --isExternal
```

Add a navigation node below an existing node

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --parentNodeId 2010 --title About --url /sites/team-s/sitepages/about.aspx
```

## Response

=== "JSON"

    ```json
    {
      "AudienceIds": null,
      "CurrentLCID": 1033,
      "Id": 2030,
      "IsDocLib": true,
      "IsExternal": true,
      "IsVisible": true,
      "ListTemplateType": 0,
      "Title": "Navigation Link",
      "Url": "https://contoso.sharepoint.com"
    }
    ```

=== "Text"

    ```text
    AudienceIds     : null
    CurrentLCID     : 1033
    Id              : 2031
    IsDocLib        : true
    IsExternal      : true
    IsVisible       : true
    ListTemplateType: 0
    Title           : Navigation Link
    Url             : https://contoso.sharepoint.com
    ```

=== "CSV"

    ```csv
    AudienceIds,CurrentLCID,Id,IsDocLib,IsExternal,IsVisible,ListTemplateType,Title,Url
    ,1033,2032,1,1,1,0,Navigation Link,https://contoso.sharepoint.com
    ```
