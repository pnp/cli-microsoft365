# pp solution publisher list

Lists publishers in a given environment.

## Usage

```sh
m365 pp solution publisher list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`--includeMicrosoftPublishers`
: Include the Microsoft Publishers

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List all publishers in a specific environment

```sh
m365 pp solution publisher list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
```

List all publishers in a specific environment as Admin

```sh
m365 pp solution publisher list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "publisherid": "00000001-0000-0000-0000-00000000005a",
        "uniquename": "Cree38e",
        "friendlyname": "CDS Default Publisher",
        "versionnumber": 1074060,
        "isreadonly": false,
        "description": null,
        "customizationprefix": "cr6c3",
        "customizationoptionvalueprefix": 43186
      }
    ]
    ```

=== "Text"

    ```text
    publisherid                           uniquename                     friendlyname
    ------------------------------------  -----------------------------  ----------------------------------
    00000001-0000-0000-0000-00000000005a  Cree38e                        CDS Default Publisher
    ```

=== "CSV"

    ```csv
    publisherid,uniquename,friendlyname
    00000001-0000-0000-0000-00000000005a,Cree38e,CDS Default Publisher
    ```

=== "Markdown"

    ```md
    # pp solution publisher list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
    
    Date: 9/1/2023

    Property | Value
    ---------|-------
    publisherid | 00000001-0000-0000-0000-00000000005a
    uniquename | Cree38e
    friendlyname | CDS Default Publisher
    versionnumber | 1074060
    isreadonly | false
    description | null
    customizationprefix | cr6c3
    customizationoptionvalueprefix | 43186
    ```
