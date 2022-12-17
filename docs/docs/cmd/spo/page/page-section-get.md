# spo page section get

Get information about the specified modern page section

## Usage

```sh
m365 spo page section get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

`-n, --pageName <pageName>`
: Name of the page to get section information of.

`-s, --section <sectionId>`
: ID of the section for which to retrieve information.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Get information about the specified section of the modern page

```sh
m365 spo page section get --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --section 1
```

## Response

=== "JSON"

    ```json
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
    ```

=== "Text"

    ```text
    columns: [{"factor":12,"order":1}]
    order  : 1
    ```

=== "CSV"

    ```csv
    order,columns
    1,"[{""factor"":12,""order"":1}]"
    ```
