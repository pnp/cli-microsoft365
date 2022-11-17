# flow list

Lists Power Automate flow in the given environment

## Usage

```sh
m365 flow list [options]
```

## Options

`-e, --environmentName <environmentName>`
: The name of the environment for which to retrieve available flows

`--asAdmin`
: Set, to list all Flows as admin. Otherwise will return only your own flows

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

By default, the `flow list` command returns only your flows. To list all flows, use the `asAdmin` option.

## Examples

List all your flows in the given environment

```sh
m365 flow list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

List all flows in the given environment

```sh
m365 flow list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
```

## Response

### Standard response

=== "JSON"

    ``` json
    [
      {
        "name": "00afcb83-df7b-4fe0-ab9c-1542a1dc66a9",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000/flows/00afcb83-df7b-4fe0-ab9c-1542a1dc66a9",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Contoso Invoicing Flow",
          "userType": "Owner",
          "state": "Started",
          "connectionReferences": {
            "shared_contoso-20invoicing-5fdd00e4805bfffb8f-5fbaee43593a7efda0": {
              "connectionName": "bd877f62e4224011aa936d706fc68902",
              "source": "Invoker",
              "id": "/providers/Microsoft.PowerApps/apis/shared_contoso-20invoicing-5fdd00e4805bfffb8f-5fbaee43593a7efda0",
              "displayName": "Contoso Invoicing",
              "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png",
              "brandColor": "#007ee5",
              "tier": "NotSpecified"
            },
            "shared_flowpush": {
              "connectionName": "shared-flowpush-d2c01136-3f7d-4449-b4f1-cb2d03a35ba8",
              "source": "Invoker",
              "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
              "displayName": "Notifications",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1599/1.0.1599.3017/flowpush/icon.png",
              "brandColor": "#FF3B30",
              "tier": "Standard"
            }
          },
          "createdTime": "2022-06-11T10:34:03.7241198Z",
          "lastModifiedTime": "2022-06-11T10:35:54.1920032Z",
          "environment": {
            "name": "Default-00000000-0000-0000-0000-000000000000",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000"
          },
          "definitionSummary": {
            "triggers": [
              {
                "type": "Request",
                "kind": "Button",
                "metadata": {
                  "operationMetadataId": "0cc0490e-e1b6-4a19-b313-f54862d64f02"
                }
              }
            ],
            "actions": [
              {
                "type": "OpenApiConnection",
                "swaggerOperationId": "ListInvoices",
                "metadata": {
                  "operationMetadataId": "d76a7b54-48bb-49a0-86b8-dd3d21b3d5ce"
                }
              },
              {
                "type": "Table",
                "metadata": {
                  "operationMetadataId": "1164ebc4-b501-46bc-bc88-cc99660f92c3"
                }
              },
              {
                "type": "OpenApiConnection",
                "swaggerOperationId": "SendEmailNotification",
                "metadata": {
                  "operationMetadataId": "9febe29f-2e36-4765-81ab-83645d28332d"
                }
              }
            ]
          },
          "creator": {
            "tenantId": "00000000-0000-0000-0000-000000000000",
            "objectId": "00000000-0000-0000-0000-000000000000",
            "userId": "00000000-0000-0000-0000-000000000000",
            "userType": "ActiveDirectory"
          },
          "provisioningMethod": "FromDefinition",
          "flowFailureAlertSubscribed": true,
          "isManaged": false
        },
        "displayName": "Contoso Invoicing Flow"
      },
      {
        "name": "15ce8985-fc2d-4043-9d8a-06a971495a99",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000/flows/15ce8985-fc2d-4043-9d8a-06a971495a99",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Recruitment Approval",
          "userType": "Owner",
          "state": "Started",
          "connectionReferences": {
            "shared_sharepointonline": {
              "connectionName": "shared-sharepointonl-1c433d0f-a030-45eb-9795-8c2585f84781",
              "source": "Embedded",
              "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1610-greyhound-localization-RelayFix/1.0.1610.3091/sharepointonline/icon.png",
              "brandColor": "#036C70",
              "tier": "Standard"
            }
          },
          "createdTime": "2022-10-17T15:25:55.1266576Z",
          "lastModifiedTime": "2022-10-17T15:25:58.8108454Z",
          "environment": {
            "name": "Default-00000000-0000-0000-0000-000000000000",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000"
          },
          "definitionSummary": {
            "triggers": [
              {
                "type": "OpenApiConnection",
                "swaggerOperationId": "GetOnUpdatedItems",
                "metadata": {
                  "operationMetadataId": "a7c43f84-ba84-4961-b4b4-a212a60cdb9d"
                }
              }
            ],
            "actions": [
              {
                "type": "If",
                "metadata": {
                  "operationMetadataId": "462baceb-7316-42fc-9436-d60716ad55b4"
                }
              },
              {
                "type": "OpenApiConnection",
                "swaggerOperationId": "PatchItem",
                "metadata": {
                  "operationMetadataId": "698b9b79-abf9-4f63-8f8b-ff584c2b8677"
                }
              },
              {
                "type": "OpenApiConnection",
                "swaggerOperationId": "PatchItem",
                "metadata": {
                  "operationMetadataId": "96f3fd44-9c7c-4521-a552-65f5f81438a0"
                }
              }
            ]
          },
          "creator": {
            "tenantId": "00000000-0000-0000-0000-000000000000",
            "objectId": "00000000-0000-0000-0000-000000000000",
            "userId": "00000000-0000-0000-0000-000000000000",
            "userType": "ActiveDirectory"
          },
          "provisioningMethod": "FromDefinition",
          "flowFailureAlertSubscribed": true,
          "isManaged": false
        },
        "displayName": "Recruitment Approval"
      },
      {
        "name": "536131e8-1dd0-4792-806d-309131261a3d",
        "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000/flows/536131e8-1dd0-4792-806d-309131261a3d",
        "type": "Microsoft.ProcessSimple/environments/flows",
        "properties": {
          "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
          "displayName": "Get Group Owners",
          "userType": "Owner",
          "state": "Suspended",
          "connectionReferences": {},
          "createdTime": "2020-01-29T07:02:43.4138109Z",
          "lastModifiedTime": "2020-01-30T06:55:31.2413475Z",
          "environment": {
            "name": "Default-00000000-0000-0000-0000-000000000000",
            "type": "Microsoft.ProcessSimple/environments",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000"
          },
          "definitionSummary": {
            "triggers": [
              {
                "type": "Request",
                "kind": "Http"
              }
            ],
            "actions": [
              {
                "type": "InitializeVariable"
              },
              {
                "type": "InitializeVariable"
              },
              {
                "type": "InitializeVariable"
              },
              {
                "type": "Http"
              }
            ]
          },
          "creator": {
            "tenantId": "00000000-0000-0000-0000-000000000000",
            "objectId": "00000000-0000-0000-0000-000000000000",
            "userId": "00000000-0000-0000-0000-000000000000",
            "userType": "ActiveDirectory"
          },
          "provisioningMethod": "FromDefinition",
          "flowFailureAlertSubscribed": true,
          "isManaged": false
        },
        "displayName": "Get Group Owners"
      }
    ]
    ```

=== "Text"

    ``` text
    name                                  displayName
    ------------------------------------  -----------------------------------------------
    00afcb83-df7b-4fe0-ab9c-1542a1dc66a9  Contoso Invoicing Flow
    15ce8985-fc2d-4043-9d8a-06a971495a99  Recruitment Approval
    536131e8-1dd0-4792-806d-309131261a3d  Get Group Owners
    ```

=== "CSV"

    ``` text
    name,displayName
    00afcb83-df7b-4fe0-ab9c-1542a1dc66a9,Contoso Invoicing Flow
    15ce8985-fc2d-4043-9d8a-06a971495a99,Recruitment Approval
    536131e8-1dd0-4792-806d-309131261a3d,Get Group Owners
    ```
