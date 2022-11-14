# pp dataverse table row list

Lists table rows for the given Dataverse table

## Usage

```sh
pp dataverse table row list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`--entitySetName [entitySetName]`
: The entity set name of the table. Specify either `entitySetName` or `tableName` but not both

`--tableName [tableName]`
: The name of the table. Specify either `entitySetName` or `tableName` but not both

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to

--8<-- "docs/cmd/_global.md"

## Examples

List all table rows for the given environment based on the entity set name

```sh
m365 pp dataverse table row list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --entitySetName "cr6c3_accounts"
```

List all table rows for the given environment based on the table name

```sh
m365 pp dataverse table row list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --tableName "cr6c3_account"
```

List all table rows for the given environment based on the entity set name as Admin

```sh
m365 pp dataverse table row list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --entitySetName "cr6c3_accounts" --asAdmin
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
