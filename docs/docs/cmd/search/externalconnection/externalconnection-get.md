# search externalconnection get

Allow the administrator to get a specific external connection for use in Microsoft Search.

## Usage

```sh
m365 search externalconnection get [options]
```

## Options

`-i, --id [id]`
: ID of the External Connection to get. Specify either `id` or `name`

`-n, --name [name]`
: Name of the External Connection to get. Specify either `id` or `name`

--8<-- "docs/cmd/_global.md"

## Examples

Get the External Connection by its id

```sh
m365 search externalconnection get --id "MyApp"
```

Get the External Connection by its name

```sh
m365 search externalconnection get --name "Test"
```

## Response

=== "JSON"

    ```json
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
    ```

=== "Text"

    ```text
    configuration: {"authorizedApps":["31359c7f-bd7e-475c-86db-fdb8c937548e"],"authorizedAppIds":["31359c7f-bd7e-475c-86db-fdb8c937548e"]}
    description  : CLI Test
    id           : CLITest
    name         : CLI-Test
    state        : draft
    ```

=== "CSV"

    ```csv
    id,name,description,state,configuration
    CLITest,CLI-Test,CLI Test,draft,"{""authorizedApps"":[""31359c7f-bd7e-475c-86db-fdb8c937548e""],""authorizedAppIds"":[""31359c7f-bd7e-475c-86db-fdb8c937548e""]}"
    ```

=== "Markdown"

    ```md
    # search externalconnection get --id "CLITest"

    Date: 2022-11-05

    ## CLI-Test (CLITest)

    Property | Value
    ---------|-------
    id | CLITest
    name | CLI-Test
    description | CLI Test
    state | draft
    ```
