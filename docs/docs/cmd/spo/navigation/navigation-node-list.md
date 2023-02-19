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
        "AudienceIds": [
          "5786b8e8-c495-4734-b345-756733960730"
        ],
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
    Id    Title            Url
    ----  ---------------  ------------------------------
    2032  Navigation Link  https://contoso.sharepoint.com
    ```

=== "CSV"

    ```csv
    Id,Title,Url
    2032,Navigation Link,https://contoso.sharepoint.com
    ```

=== "Markdown"

    ```md
    # spo navigation node list --webUrl "https://contoso.sharepoint.com" --location "QuickLaunch"

    Date: 27/1/2023

    ## Home (1031)

    Property | Value
    ---------|-------
    AudienceIds | ["5786b8e8-c495-4734-b345-756733960730"]
    CurrentLCID | 1033
    Id | 2032
    IsDocLib | true
    IsExternal | false
    IsVisible | true
    ListTemplateType | 0
    Title | Navigation Link
    Url | https://contoso.sharepoint.com
    ```
