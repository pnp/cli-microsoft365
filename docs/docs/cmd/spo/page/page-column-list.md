# spo page column list

Lists columns in the specific section of a modern page

## Usage

```sh
m365 spo page column list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located.

`-n, --pageName <pageName>`
: Name of the page to list columns of.

`-s, --section <sectionId>`
: ID of the section for which to list columns.

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List columns in the first section of a modern page

```sh
m365 spo page column list --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --section 1
```


## Response

=== "JSON"

    ```json
    [
      {
        "factor": 12,
        "order": 1,
        "dataVersion": "1.0",
        "jsonData": "&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionFactor&quot;&#58;12,&quot;sectionIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1&#125;&#125;",
        "controls": 1
      }
    ]
    ```

=== "Text"

    ```text
    controls: 1
    factor  : 12
    order   : 1
    ```

=== "CSV"

    ```csv
    factor,order,controls
    12,1,1
    ```
