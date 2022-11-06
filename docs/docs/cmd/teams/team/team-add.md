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

## Response

### Standard response

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
    
### `wait` response

When we make use of the option `wait` the response will differ. 

=== "JSON"

    ``` json
    {
      "id": "d592059d-100f-48c6-8a91-b68eec00ecec",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2022-11-04T12:46:47Z",
      "creationOptions": [
        "Team",
        "ExchangeProvisioningFlags:3552"
      ],
      "description": "Architecture Discussion",
      "displayName": "Architecture",
      "expirationDateTime": null,
      "groupTypes": [
        "Unified"
      ],
      "isAssignableToRole": null,
      "mail": "Architecture@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "Architecture",
      "membershipRule": null,
      "membershipRuleProcessingState": null,
      "onPremisesDomainName": null,
      "onPremisesLastSyncDateTime": null,
      "onPremisesNetBiosName": null,
      "onPremisesSamAccountName": null,
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "preferredLanguage": null,
      "proxyAddresses": [
        "SMTP:Architecture@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2022-11-04T12:46:47Z",
      "resourceBehaviorOptions": [
        "HideGroupInOutlook",
        "SubscribeMembersToCalendarEventsDisabled",
        "WelcomeEmailDisabled"
      ],
      "resourceProvisioningOptions": [
        "Team"
      ],
      "securityEnabled": false,
      "securityIdentifier": "S-1-12-1-3583116701-1220939791-2394329482-3974889708",
      "theme": null,
      "visibility": "Public",
      "onPremisesProvisioningErrors": []
    }
    ```

=== "Text"

    ``` text
    classification               : null
    createdDateTime              : 2022-11-04T12:47:57Z
    creationOptions              : ["Team","ExchangeProvisioningFlags:3552"]
    deletedDateTime              : null
    description                  : Architecture Discussion
    displayName                  : Architecture
    expirationDateTime           : null
    groupTypes                   : ["Unified"]
    id                           : 29c242bb-a96f-470a-b280-d63154f5446f
    isAssignableToRole           : null
    mail                         : Architecture@contoso.onmicrosoft.com
    mailEnabled                  : true
    mailNickname                 : Architecture
    membershipRule               : null
    membershipRuleProcessingState: null
    onPremisesDomainName         : null
    onPremisesLastSyncDateTime   : null
    onPremisesNetBiosName        : null
    onPremisesProvisioningErrors : []
    onPremisesSamAccountName     : null
    onPremisesSecurityIdentifier : null
    onPremisesSyncEnabled        : null
    preferredDataLocation        : null
    preferredLanguage            : null
    proxyAddresses               : ["SMTP:Architecture@contoso.onmicrosoft.com"]
    renewedDateTime              : 2022-11-04T12:47:57Z
    resourceBehaviorOptions      : ["HideGroupInOutlook","SubscribeMembersToCalendarEventsDisabled","WelcomeEmailDisabled"]
    resourceProvisioningOptions  : ["Team"]
    securityEnabled              : false
    securityIdentifier           : S-1-12-1-700596923-1191881071-836141234-1866790228
    theme                        : null
    visibility                   : Public
    ```

=== "CSV"

    ``` text
    id,deletedDateTime,classification,createdDateTime,creationOptions,description,displayName,expirationDateTime,groupTypes,isAssignableToRole,mail,mailEnabled,mailNickname,membershipRule,membershipRuleProcessingState,onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,preferredDataLocation,preferredLanguage,proxyAddresses,renewedDateTime,resourceBehaviorOptions,resourceProvisioningOptions,securityEnabled,securityIdentifier,theme,visibility,onPremisesProvisioningErrors
    bb57868a-e82e-470b-85aa-8a86942a5bf8,,,2022-11-04T12:51:35Z,"[""Team"",""ExchangeProvisioningFlags:3552""]",Architecture Discussion,Architecture,,"[""Unified""]",,Architecture@contoso.onmicrosoft.com,1,TeamName,,,,,,,,,,,"[""SMTP:Architecture@contoso.onmicrosoft.com""]",2022-11-04T12:51:35Z,"[""HideGroupInOutlook"",""SubscribeMembersToCalendarEventsDisabled"",""WelcomeEmailDisabled""]","[""Team""]",,S-1-12-1-3143075466-1191962670-2257234565-4166724244,,Public,[]
    ```

## More information

- Get started with Teams templates: [https://docs.microsoft.com/MicrosoftTeams/get-started-with-teams-templates](https://docs.microsoft.com/MicrosoftTeams/get-started-with-teams-templates)
- group resource type: [https://docs.microsoft.com/graph/api/resources/group?view=graph-rest-1.0](https://docs.microsoft.com/graph/api/resources/group?view=graph-rest-1.0)
- teamsAsyncOperation resource type: [https://docs.microsoft.com/graph/api/resources/teamsasyncoperation?view=graph-rest-1.0](https://docs.microsoft.com/graph/api/resources/teamsasyncoperation?view=graph-rest-1.0)
