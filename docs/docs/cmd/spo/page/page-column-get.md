# spo page column get

Get information about a specific column of a modern page

## Usage

```sh
m365 spo page column get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

`-n, --pageName <pageName>`
: Name of the page to get column information of.

`-s, --section <section>`
: ID of the section where the column is located.

`-c, --column <column>`
: ID of the column for which to retrieve more information.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Get information about the first column in the first section of a modern page

```sh
m365 spo page column get --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --section 1 --column 1
```

## Response

=== "JSON"

    ```json
    {
      "factor": 12,
      "order": 1,
      "dataVersion": "1.0",
      "jsonData": "&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;",
      "controls": [
        {
          "controlType": 3,
          "dataVersion": "1.0",
          "order": 1,
          "id": "7558d804-0334-49ca-b14a-53870cf6caae",
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
            "emphasis": {}
          },
          "title": "Bing Maps",
          "description": "Display a location on a map using Bing Maps.",
          "propertieJson": {
            "pushPins": [],
            "maxNumberOfPushPins": 1,
            "shouldShowPushPinTitle": true,
            "zoomLevel": 12,
            "mapType": "road"
          },
          "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
          "htmlProperties": "",
          "serverProcessedContent": null,
          "canvasDataVersion": "1.0"
        }
      ]
    }
    ```

=== "Text"

    ```text
    controls: 7558d804-0334-49ca-b14a-53870cf6caae (Bing Maps)
    factor  : 12
    order   : 1
    ```

=== "CSV"

    ```csv
    factor,order,controls
    12,1,7558d804-0334-49ca-b14a-53870cf6caae (Bing Maps)
    ```

=== "Markdown"

    ```md
    # spo page column get --webUrl "https://contoso.sharepoint.com/sites/team-a" --pageName "home.aspx" --section "1" --column "1"

    Date: 5/1/2023

    Property | Value
    ---------|-------
    factor | 6
    order | 1
    dataVersion | 1.0
    jsonData | &#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;6,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;
    ```
