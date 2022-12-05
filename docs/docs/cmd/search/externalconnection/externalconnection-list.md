# search externalconnection list

Lists external connections defined in Microsoft Search

## Usage

```sh
m365 search externalconnection list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

List external connections defined in Microsoft Search

```sh
m365 search externalconnection list
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "CLITest",
        "name": "CLI-Test",
        "description": "CLI Test",
        "state": "draft",
        "configuration": {
          "authorizedApps": [
            "31359c7f-bd7e-475c-86db-fdb8c937548e"
          ],
          "authorizedAppIds": [
            "31359c7f-bd7e-475c-86db-fdb8c937548e"
          ]
        }
      }
    ]
    ```

=== "Text"

    ```text
    id       name      state
    -------  --------  -----
    CLITest  CLI-Test  draft
    ```

=== "CSV"

    ```csv
    id,name,state
    CLITest,CLI-Test,draft
    ```
