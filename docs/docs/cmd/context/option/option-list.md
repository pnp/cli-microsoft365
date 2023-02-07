# context option list

List all options added to the context

## Usage

```sh
m365 context option list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

List all options added to the context

```sh
m365 context option list
```

## Response

The responses below are an example. The output may differ based on the contents of the context file

=== "JSON"

    ```json
    {
      "url": "https://contoso.sharepoint.com",
      "list": "list name"
    }
    ```

=== "Text"

    ```text
    list: list name
    url : https://contoso.sharepoint.com
    ```

=== "CSV"

    ```csv
    url,list
    https://contoso.sharepoint.com,list name
    ```

=== "Markdown"

    ```md
    # context option list

    Date: 7/2/2023

    ## https://contoso.sharepoint.com

    Property | Value
    ---------|-------
    url | https://contoso.sharepoint.com
    list | list name
    ```
