# spo page control list

Lists controls on the specific modern page

## Usage

```sh
m365 spo page control list [options]
```

## Options

`-n, --pageName <pageName>`
: Name of the page to list controls of.

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

List controls on the modern page

```sh
m365 spo page control list --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx
```


## Response

=== "JSON"

    ```json
    [
      {
        "id": "7558d804-0334-49ca-b14a-53870cf6caae",
        "type": "Client-side web part",
        "title": "Bing Maps",
        "controlType": 3,
        "order": 1,
        "controlData": {
          "controlType": 3,
          "displayMode": 2,
          "id": "7558d804-0334-49ca-b14a-53870cf6caae",
          "position": {
            "controlIndex": 1,
            "sectionIndex": 1,
            "zoneIndex": 1,
            "sectionFactor": 12,
            "layoutIndex": 1
          },
          "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
          "emphasis": {},
          "webPartData": {
            "dataVersion": "1.0",
            "description": "Display a location on a map using Bing Maps.",
            "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
            "instanceId": "7558d804-0334-49ca-b14a-53870cf6caae",
            "properties": {
              "pushPins": [],
              "maxNumberOfPushPins": 1,
              "shouldShowPushPinTitle": true,
              "zoomLevel": 12,
              "mapType": "road"
            },
            "title": "Bing Maps",
            "serverProcessedContent": {
              "htmlStrings": {},
              "searchablePlainTexts": {},
              "imageSources": {},
              "links": {}
            }
          }
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                    title      type
    ------------------------------------  ---------  ---------------------
    7558d804-0334-49ca-b14a-53870cf6caae  Bing Maps  Client-side web part
    ```

=== "CSV"

    ```csv
    id,type,title
    7558d804-0334-49ca-b14a-53870cf6caae,Client-side web part,Bing Maps
    ```

=== "Markdown"

    ```md
    # spo page control list --webUrl "https://contoso.sharepoint.com/sites/team-a" --pageName "home.aspx"

    Date: 5/1/2023

    ## Bing Maps (f85f8dfa-9052-4be8-8954-8cdafe811b97)

    Property | Value
    ---------|-------
    id | f85f8dfa-9052-4be8-8954-8cdafe811b97
    type | Client-side web part
    title | Bing Maps
    controlType | 3
    order | 1
    ```
