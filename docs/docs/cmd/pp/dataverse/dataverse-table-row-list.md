# pp dataverse table row list

Lists dataverse table rows in a given environment

## Usage

```sh
pp dataverse table row list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment to list all table rows for

`-n, --name <name>`
: The name of the table. Note that this is the logical name in the plural

`-a, --asAdmin`
: Set, to retrieve the dataverse table rows as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Examples

List all table rows for the given environment

```sh
m365 pp dataverse table row list -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "cr6c3_accounts"
```

List all table rows for the given environment as Admin

```sh
m365 pp dataverse table row list -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "cr6c3_accounts" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "cr6c3_accountsid": "95c80273-3764-ed11-9561-000d3a4bbea4",
        "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
        "statecode": 0,
        "statuscode": 1,
        "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "modifiedon": "2022-11-14T16:14:45Z",
        "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "versionnumber": 1413873,
        "createdon": "2022-11-14T16:14:45Z",
        "cr6c3_name": "Column1 value",
        "overriddencreatedon": null,
        "importsequencenumber": null,
        "_modifiedonbehalfby_value": null,
        "utcconversiontimezonecode": null,
        "_createdonbehalfby_value": null,
        "_owningteam_value": null,
        "timezoneruleversionnumber": null
      }
    ]
    ```

=== "Text"

    ```text
    cr6c3_accountsid                       _owningbusinessunit_value             statecode  statuscode  _createdby_value                      _ownerid_value                        modifiedon            _owninguser_value                     _modifiedby_value                     versionnumber  createdon             cr6c3_name      overriddencreatedon  importsequencenumber  _modifiedonbehalfby_value  utcconversiontimezonecode  _createdonbehalfby_value  _owningteam_value  timezoneruleversionnumber
    ------------------------------------  ------------------------------------  ---------  ----------  ------------------------------------  ------------------------------------  --------------------  ------------------------------------  ------------------------------------  -------------  --------------------  ----------      -------------------  --------------------  -------------------------  -------------------------  ------------------------  -----------------  -------------------------
    95c80273-3764-ed11-9561-000d3a4bbea4  6da087c1-1c4d-ed11-bba1-000d3a2caf7f  0          1           5fa787c1-1c4d-ed11-bba1-000d3a2caf7f  5fa787c1-1c4d-ed11-bba1-000d3a2caf7f  2022-11-14T16:14:45Z  5fa787c1-1c4d-ed11-bba1-000d3a2caf7f  5fa787c1-1c4d-ed11-bba1-000d3a2caf7f  1413873        2022-11-14T16:14:45Z  Column1 value   null                 null                  null                       null                       null                      null               null
    ```

=== "CSV"

    ```csv
    cr6c3_accountsid,_owningbusinessunit_value,statecode,statuscode,_createdby_value,_ownerid_value,modifiedon,_owninguser_value,_modifiedby_value,versionnumber,createdon,cr6c3_name,overriddencreatedon,importsequencenumber,_modifiedonbehalfby_value,utcconversiontimezonecode,_createdonbehalfby_value,_owningteam_value,timezoneruleversionnumber
    95c80273-3764-ed11-9561-000d3a4bbea4,6da087c1-1c4d-ed11-bba1-000d3a2caf7f,0,1,5fa787c1-1c4d-ed11-bba1-000d3a2caf7f,5fa787c1-1c4d-ed11-bba1-000d3a2caf7f,2022-11-14T16:14:45Z,5fa787c1-1c4d-ed11-bba1-000d3a2caf7f,5fa787c1-1c4d-ed11-bba1-000d3a2caf7f,1413873,2022-11-14T16:14:45Z,Column1 value,,,,,,,
    ```
