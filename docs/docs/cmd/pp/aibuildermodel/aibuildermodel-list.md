# pp aibuildermodel list

List available AI builder models in the specified Power Platform environment

## Usage

```sh
pp aibuildermodel list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to

--8<-- "docs/cmd/_global.md"

## Examples

List all AI Builder models in a specific environment

```sh
m365 pp aibuildermodel list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
```

List all AI Builder models in a specific environment as admin

```sh
m365 pp aibuildermodel list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "statecode": 0,
        "_msdyn_templateid_value": "10707e4e-1d56-e911-8194-000d3a6cd5a5",
        "msdyn_modelcreationcontext": "{}",
        "createdon": "2022-11-29T11:58:45Z",
        "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "modifiedon": "2022-11-29T11:58:45Z",
        "msdyn_sharewithorganizationoncreate": false,
        "msdyn_aimodelidunique": "b0328b67-47e2-4202-8189-e617ec9a88bd",
        "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
        "ismanaged": false,
        "versionnumber": 1458121,
        "msdyn_name": "Document Processing 11/29/2022, 12:58:43 PM",
        "introducedversion": "1.0",
        "statuscode": 0,
        "_modifiedby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "overwritetime": "1900-01-01T00:00:00Z",
        "componentstate": 0,
        "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
        "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "msdyn_aimodelid": "08ffffbe-ec1c-4e64-b64b-dd1db926c613",
        "_msdyn_activerunconfigurationid_value": null,
        "overriddencreatedon": null,
        "_msdyn_retrainworkflowid_value": null,
        "importsequencenumber": null,
        "_msdyn_scheduleinferenceworkflowid_value": null,
        "_modifiedonbehalfby_value": null,
        "utcconversiontimezonecode": null,
        "_createdonbehalfby_value": null,
        "_owningteam_value": null,
        "timezoneruleversionnumber": null,
        "iscustomizable": {
          "Value": true,
          "CanBeChanged": true,
          "ManagedPropertyLogicalName": "iscustomizableanddeletable"
        }
      }
    ]
    ```

=== "Text"

    ```text
    createdon             modifiedon            msdyn_aimodelid                       msdyn_name
    --------------------  --------------------  ------------------------------------  -------------------------------------------
    2022-10-25T14:44:48Z  2022-10-25T14:44:48Z  08ffffbe-ec1c-4e64-b64b-dd1db926c613  Document Processing 11/29/2022, 12:58:43 PM
    ```

=== "CSV"

    ```csv
    msdyn_name,msdyn_aimodelid,createdon,modifiedon
    "Document Processing 11/29/2022, 12:58:43 PM",08ffffbe-ec1c-4e64-b64b-dd1db926c613,2022-11-29T11:58:45Z,2022-11-29T11:58:45Z
    ```
