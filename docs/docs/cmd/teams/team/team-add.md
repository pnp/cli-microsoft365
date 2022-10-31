# teams team add

Adds a new Microsoft Teams team

## Usage

```sh
m365 teams team add [options]
```

## Options

`-n, --name [name]`
: Display name for the Microsoft Teams team. Required if `template` not supplied

`-d, --description [description]`
: Description for the Microsoft Teams team. Required if `template` not supplied

`--template [template]`
: Template to use to create the team. If `name` or `description` are supplied, these take precedence over the template values

`--wait`
: Wait for the team to be provisioned before completing the command

--8<-- "docs/cmd/_global.md"

## Remarks

If you want to add a Team to an existing Microsoft 365 Group use the [aad o365group teamify](../../aad/o365group/o365group-teamify.md) command instead.

This command will return different responses based on the presence of the `--wait` option. If present, the command will return a `group` resource in the response. If not present, the command will return a `teamsAsyncOperation` resource in the response.

## Examples

Add a new Microsoft Teams team

```sh
m365 teams team add --name "Architecture" --description "Architecture Discussion"
```

Add a new Microsoft Teams team using a template from a file

```sh
m365 teams team add --name "Architecture" --description "Architecture Discussion" --template @template.json
```

Add a new Microsoft Teams team using a template and wait for the team to be provisioned

```sh
m365 teams team add --name "Architecture" --description "Architecture Discussion" --template @template.json --wait
```

## More information

- Get started with Teams templates: [https://docs.microsoft.com/MicrosoftTeams/get-started-with-teams-templates](https://docs.microsoft.com/MicrosoftTeams/get-started-with-teams-templates)
- group resource type: [https://docs.microsoft.com/graph/api/resources/group?view=graph-rest-1.0](https://docs.microsoft.com/graph/api/resources/group?view=graph-rest-1.0)
- teamsAsyncOperation resource type: [https://docs.microsoft.com/graph/api/resources/teamsasyncoperation?view=graph-rest-1.0](https://docs.microsoft.com/graph/api/resources/teamsasyncoperation?view=graph-rest-1.0)

## Response

=== "JSON"

    ``` json
    {
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('a40210cd-0060-4b91-aaa1-a44e0853d979')/operations/$entity",
      "id": "d708ecb3-3325-4f6e-a0f7-2f982901b856",
      "operationType": "createTeam",
      "createdDateTime": "2022-10-31T12:50:44.0819314Z",
      "status": "notStarted",
      "lastActionDateTime": "2022-10-31T12:50:44.0819314Z",
      "attemptsCount": 1,
      "targetResourceId": "a40210cd-0060-4b91-aaa1-a44e0853d979",
      "targetResourceLocation": "/teams('a40210cd-0060-4b91-aaa1-a44e0853d979')",
      "Value": "{\"apps\":[],\"channels\":[],\"WorkflowId\":\"westeurope.0837160b-803e-4279-9f2c-a5cc46ffc748\"}",
      "error": null
    }
    ```

=== "Text"

    ``` text
    @odata.context        : https://graph.microsoft.com/v1.0/$metadata#teams('6d9e3e6b-88a2-492b-985a-477bc760bd6b')/operations/$entity
    Value                 : {"apps":[],"channels":[],"WorkflowId":"FranceCentral.7bfcab39-032e-4a1a-b4f0-4e5fb035b5a1"}
    attemptsCount         : 1
    createdDateTime       : 2022-10-31T12:51:22.8337964Z
    error                 : null
    id                    : 6af66f9c-f73b-42a6-87b8-216dee12f40b
    lastActionDateTime    : 2022-10-31T12:51:22.8337964Z
    operationType         : createTeam
    status                : notStarted
    targetResourceId      : 6d9e3e6b-88a2-492b-985a-477bc760bd6b
    targetResourceLocation: /teams('6d9e3e6b-88a2-492b-985a-477bc760bd6b')
    ```

=== "CSV"

    ``` text
    @odata.context,id,operationType,createdDateTime,status,lastActionDateTime,attemptsCount,targetResourceId,targetResourceLocation,Value,error
    https://graph.microsoft.com/v1.0/$metadata#teams('40d5758d-5ad9-406d-88ab-0a78992ffbab')/operations/$entity,65778567-595d-4543-bb21-f8d62c678c8e,createTeam,2022-10-31T12:57:42.4956529Z,notStarted,2022-10-31T12:57:42.4956529Z,1,40d5758d-5ad9-406d-88ab-0a78992ffbab,/teams('40d5758d-5ad9-406d-88ab-0a78992ffbab'),"{""apps"":[],""channels"":[],""WorkflowId"":""northeurope.d0475d7e-7461-4dd5-ae1e-0cfa9e692412""}",
    ```
