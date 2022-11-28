# spo propertybag list

Gets property bag values

## Usage

```sh
m365 spo propertybag list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site from which the property bag value should be retrieved.

`-f, --folder [folder]`
: Site-relative URL of the folder from which to retrieve property bag value. Case-sensitive.

--8<-- "docs/cmd/_global.md"

## Examples

Return property bag values located in the given site

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test
```

Return property bag values located in the given site root folder

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder /
```

Return property bag values located in the given site document library

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents'
```

Return property bag values located in folder in the given site document library

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents/MyFolder'
```

Return property bag values located in the given site list

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder /Lists/MyList
```

## Response

=== "JSON"

    ```json
    [
      {
        "key": "vti_approvallevels",
        "value": "Approved Rejected Pending\\ Review"
      }
    ]
    ```

=== "Text"

    ```text
    key                   value                                                                             
    --------------------  -----------------------------------
    vti_approvallevels     Approved Rejected Pending\ Review
    ```

=== "CSV"

    ```csv
    key,value
    vti_approvallevels,Approved Rejected Pending\ Review
    ```
