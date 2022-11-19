# pp chatbot get

Gets a specific Microsoft Power Platform chatbot in the specified Power Platform environment

## Usage

```sh
pp chatbot get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id [id]`
: The id of the chatbot. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the chatbot. Specify either `id` or `name` but not both.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific chatbot in a specific environment based on name

```sh
m365 pp chatbot get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Chatbot"
```

Get a specific chatbot in a specific environment based on name as admin

```sh
m365 pp chatbot get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --name "CLI 365 Chatbot" --asAdmin
```

Get a specific chatbot in a specific environment based on id

```sh
m365 pp chatbot get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893"
```

Get a specific chatbot in a specific environment based on id as admin

```sh
m365 pp chatbot get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893" --asAdmin
```

## Response

=== "JSON"

    ```json
    {
      "authenticationtrigger": 0,
      "_owningbusinessunit_value": "6da087c1-1c4d-ed11-bba1-000d3a2caf7f",
      "statuscode": 1,
      "createdon": "2022-11-19T10:42:22Z",
      "statecode": 0,
      "schemaname": "new_bot_23f5f58697fd43d595eb451c9797a53d",
      "_ownerid_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
      "overwritetime": "1900-01-01T00:00:00Z",
      "name": "CLI 365 Chatbot",
      "solutionid": "fd140aae-4df4-11dd-bd17-0019b9312238",
      "ismanaged": false,
      "versionnumber": 1421457,
      "language": 1033,
      "_modifiedby_value": "5f91d7a7-5f46-494a-80fa-5c18b0221351",
      "_modifiedonbehalfby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
      "modifiedon": "2022-11-19T10:42:24Z",
      "componentstate": 0,
      "botid": "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893",
      "_createdby_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
      "componentidunique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23",
      "authenticationmode": 1,
      "_owninguser_value": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
      "accesscontrolpolicy": 0,
      "runtimeprovider": 0,
      "_publishedby_value": "John Doe",
      "authenticationconfiguration": null,
      "authorizedsecuritygroupids": null,
      "overriddencreatedon": null,
      "applicationmanifestinformation": null,
      "importsequencenumber": null,
      "synchronizationstatus": null,
      "template": null,
      "_providerconnectionreferenceid_value": null,
      "configuration": null,
      "utcconversiontimezonecode": null,
      "publishedon": "2022-11-19T10:43:24Z",
      "_createdonbehalfby_value": null,
      "iconbase64": null,
      "supportedlanguages": null,
      "_owningteam_value": null,
      "timezoneruleversionnumber": null,
      "iscustomizable": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "iscustomizableanddeletable"
      }
    }
    ```

=== "Text"

    ```text
    botid      : 3a081d91-5ea8-40a7-8ac9-abbaa3fcb893
    createdon  : 2022-11-19T10:42:22Z
    modifiedon : 2022-11-19T10:42:24Z
    name       : CLI 365 Chatbot
    publishedon: 2022-11-19T10:43:24Z
    ```

=== "CSV"

    ```csv
    name,botid,publishedon,createdon,modifiedon
    CLI 365 Chatbot,3a081d91-5ea8-40a7-8ac9-abbaa3fcb893,2022-11-19T10:43:24Z,2022-11-19T10:42:22Z,2022-11-19T10:42:24Z
    ```
