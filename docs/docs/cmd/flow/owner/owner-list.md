# flow owner list

Lists all owners of a Power Automate flow

## Usage

```sh
m365 flow owner list [options]
```

## Options

`-e, --environmentName <environmentName>`
: The name of the environment.

`-n, --name <name>`
: The name of the Power Automate flow.

`--asAdmin`
: Run the command as admin.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the owners by the name of the Power Automate flow within a specified environment

```sh
m365 flow owner list --environmentName Default-c5a5d746-3520-453f-8a69-780f8e44917e --name 72f2be4a-78c1-4220-a048-dbf557296a72
```

Gets the owners by the name of the Power Automate flow within a specified environment with admin permissions

```sh
m365 flow owner list --environmentName Default-c5a5d746-3520-453f-8a69-780f8e44917e --name 72f2be4a-78c1-4220-a048-dbf557296a72 --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "name": "fe36f75e-c103-410b-a18a-2bf6df06ac3a",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-c5a5d746-3520-453f-8a69-780f8e44917e/flows/72f2be4a-78c1-4220-a048-dbf557296a72/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a",
        "type": "/providers/Microsoft.ProcessSimple/environments/flows/permissions",
        "properties": {
          "roleName": "Owner",
          "permissionType": "Principal",
          "principal": {
            "id": "fe36f75e-c103-410b-a18a-2bf6df06ac3a",
            "type": "User"
          }
        }
      }
    ]
    ```

=== "Text"

    ```text
    roleName  id                                    type
    --------  ------------------------------------  ----
    Owner     fe36f75e-c103-410b-a18a-2bf6df06ac3a  User
    ```

=== "CSV"

    ```csv
    roleName,id,type
    Owner,fe36f75e-c103-410b-a18a-2bf6df06ac3a,User
    ```

=== "Markdown"

    ```md
    # flow owner list --environmentName "Default-c5a5d746-3520-453f-8a69-780f8e44917e" --name "72f2be4a-78c1-4220-a048-dbf557296a72" --debug "true"

    Date: 25/02/2023

    ## fe36f75e-c103-410b-a18a-2bf6df06ac3a (/providers/Microsoft.ProcessSimple/environments/Default-c5a5d746-3520-453f-8a69-780f8e44917e/flows/72f2be4a-78c1-4220-a048-dbf557296a72/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a)

    Property | Value
    ---------|-------
    name | fe36f75e-c103-410b-a18a-2bf6df06ac3a
    id | /providers/Microsoft.ProcessSimple/environments/Default-c5a5d746-3520-453f-8a69-780f8e44917e/flows/72f2be4a-78c1-4220-a048-dbf557296a72/permissions/fe36f75e-c103-410b-a18a-2bf6df06ac3a
    type | /providers/Microsoft.ProcessSimple/environments/flows/permissions
    properties | {"roleName":"Owner","permissionType":"Principal","principal":{"id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a","type":"User"}}
    ```
