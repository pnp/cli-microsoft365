# spo page section list

List sections in the specific modern page

## Usage

```sh
m365 spo page section list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

`-n, --pageName <pageName>`
: Name of the page to list sections of.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List sections of a modern page

```sh
m365 spo page section list --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx
```

## Response

=== "JSON"

    ```json
    [
      {
        "order": 1,
        "columns": [
          {
            "factor": 12,
            "order": 1,
            "dataVersion": "1.0",
            "jsonData": "&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;"
          }
        ]
      }
    ]
    ```

=== "Text"

    ```text
    order  columns
    -----  -------
    1      1
    ```

=== "CSV"

    ```csv
    order,columns
    1,1
    ```

=== "Markdown"

    ```md
    # spo page section list --webUrl "https://contoso.sharepoint.com/sites/team-a" --pageName "home.aspx"

    Date: 5/3/2023

    Property | Value
    ---------|-------
    order | 1

    ```
