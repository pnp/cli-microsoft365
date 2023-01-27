# spo navigation node list

Lists nodes from the specified site navigation

## Usage

```sh
m365 spo navigation node list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site for which to retrieve navigation

`-l, --location <location>`
: Navigation type to retrieve. Available options: `QuickLaunch,TopNavigationBar`

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve nodes from the top navigation

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar
```

Retrieve nodes from the quick launch

```sh
m365 spo navigation node list --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch
```

## Response

=== "JSON"

    ```json
    [
      {
        "AudienceIds": null,
        "CurrentLCID": 1033,
        "Id": 2032,
        "IsDocLib": true,
        "IsExternal": true,
        "IsVisible": true,
        "ListTemplateType": 0,
        "Title": "Navigation Link",
        "Url": "https://contoso.sharepoint.com"
      }
    ]
    ```

=== "Text"

    ```text
    AudienceIds  CurrentLCID  Id    IsDocLib  IsExternal  IsVisible  ListTemplateType  Title              Url
    -----------  -----------  ----  --------  ----------  ---------  ----------------  -----------------  ---------------------------
    null         1033         2032  true      false       true       0                 Navigation Link    https://contoso.sharepoint.com
    ```

=== "CSV"

    ```csv
    AudienceIds,CurrentLCID,Id,IsDocLib,IsExternal,IsVisible,ListTemplateType,Title,Url
    ,1033,2032,1,,1,0,Navigation Link,https://contoso.sharepoint.com
    ```

=== "Markdown"

    ```md
    # spo navigation node list --webUrl "https://contoso.sharepoint.com" --location "QuickLaunch"

    Date: 27/1/2023

    ## Home (1031)

    Property | Value
    ---------|-------
    AudienceIds | null
    CurrentLCID | 1033
    Id | 2032
    IsDocLib | true
    IsExternal | false
    IsVisible | true
    ListTemplateType | 0
    Title | Navigation Link
    Url | https://contoso.sharepoint.com
    ```
