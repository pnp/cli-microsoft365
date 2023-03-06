# spo navigation node add

Adds a navigation node to the specified site navigation

## Usage

```sh
m365 spo navigation node add [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site to which navigation should be modified.

`-l, --location [location]`
: Navigation type where the node should be added. Available options: `QuickLaunch`, `TopNavigationBar`. Specify either `location` or `parentNodeId` but not both.

`-t, --title <title>`
: Navigation node title.

`--url [url]`
: Navigation node URL. When not specified a linkless label will be created

`--parentNodeId [parentNodeId]`
: ID of the node below which the node should be added. Specify either `location` or `parentNodeId` but not both.

`--isExternal`
: Set, if the navigation node points to an external URL.

`--audienceIds [audienceIds]`
: Comma separated list of group IDs that will be used for audience targeting. The limit is 10 ids per navigation node.

--8<-- "docs/cmd/_global.md"

## Remarks

To enable/disable audience targeting for the navigation bar, use the [`spo web set`](../web/web-set.md) command.

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

Add a navigation node to the top navigation which is audience targetted

```sh
m365 spo navigation node add --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar --title About --url /sites/team-s/sitepages/about.aspx --audienceIds "7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a"
```

## Response

=== "JSON"

    ```json
    {
      "AudienceIds": [
        "7aa4a1ca-4035-4f2f-bac7-7beada59b5ba"
      ],
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
    AudienceIds     : ["7aa4a1ca-4035-4f2f-bac7-7beada59b5ba"]
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
    "[""7aa4a1ca-4035-4f2f-bac7-7beada59b5ba""]",1033,2032,1,1,1,0,Navigation Link,https://contoso.sharepoint.com
    ```

=== "Markdown"

    ```md
    # spo navigation node get --webUrl "https://contoso.sharepoint.com/sites/team-a" --location "TopNavigationBar" --title "Navigation Link" --url "https://contoso.sharepoint.com"

    Date: 2/20/2023

    ## Navigation Link (2030)

    Property | Value
    ---------|-------
    AudienceIds | ["7aa4a1ca-4035-4f2f-bac7-7beada59b5ba"]
    CurrentLCID | 1033
    Id | 2030
    IsDocLib | true
    IsExternal | false
    IsVisible | true
    ListTemplateType | 0
    Title | Navigation Link
    Url | https://contoso.sharepoint.com
    ```
