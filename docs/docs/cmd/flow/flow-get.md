# flow get

Gets information about the specified Power Automate flow

## Usage

```sh
m365 flow get [options]
```

## Options

`-n, --name <name>`
: The name of the Power Automate flow to get information about

`-e, --environmentName <environmentName>`
: The name of the environment for which to retrieve available flows

`--asAdmin`
: Set, to retrieve the Flow as admin

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

By default, the command will try to retrieve Power Automate flows you own. If you want to retrieve a flow owned by another user, use the `asAdmin` flag.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

If the Power Automate flow with the name you specified doesn't exist, you will get the `The caller with object id 'abc' does not have permission for connection 'xyz' under Api 'shared_logicflows'.` error. If you try to retrieve a non-existing flow as admin, you will get the `Could not find flow 'xyz'.` error.

## Examples

Get information about the specified Power Automate flow owned by the currently signed-in user

```sh
m365 flow get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d
```

Get information about the specified Power Automate flow owned by another user

```sh
m365 flow get --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
```

## Response

=== "JSON"

    ```json
    {
      "name": "ca76d7b8-3b76-4050-8c03-9fb310ad172f",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000/flows/ca76d7b8-3b76-4050-8c03-9fb310ad172f",
      "type": "Microsoft.ProcessSimple/environments/flows",
      "properties": {
        "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
        "displayName": "My Flow",
        "userType": "Owner",
        "definition": {
          "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
          "contentVersion": "1.0.0.0",
          "parameters": {
            "$connections": {
              "defaultValue": {},
              "type": "Object"
            },
            "$authentication": {
              "defaultValue": {},
              "type": "SecureObject"
            }
          },
          "triggers": {
            "When_a_file_is_created_in_a_folder": {
              "recurrence": {
                "frequency": "Minute",
                "interval": 1
              },
              "metadata": {
                "%252fComputer%2bVision%2bDemo": "/My Flow Demo",
                "operationMetadataId": "2e0aa276-5536-4f67-a7b7-c1a37a91b8ca",
                "%252fShared%2bDocuments": "/Shared Documents"
              },
              "type": "OpenApiConnection",
              "inputs": {
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                  "connectionName": "shared_sharepointonline",
                  "operationId": "OnNewFile"
                },
                "parameters": {
                  "dataset": "https://contoso.sharepoint.com/sites/contosoHome",
                  "folderId": "%252fShared%2bDocuments",
                  "inferContentType": true
                },
                "authentication": "@parameters('$authentication')"
              }
            }
          },
          "actions": {
            "Analyze_Image": {
              "runAfter": {
                "Initialize_variable": [
                  "Succeeded"
                ]
              },
              "metadata": {
                "operationMetadataId": "82b95160-9c96-4507-ad8c-c688ea95c2b1"
              },
              "type": "OpenApiConnection",
              "inputs": {
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_cognitiveservicescomputervision",
                  "connectionName": "shared_cognitiveservicescomputervision",
                  "operationId": "AnalyzeImageV2"
                },
                "parameters": {
                  "format": "Image Content",
                  "Image": "@triggerOutputs()?['body']"
                },
                "authentication": "@parameters('$authentication')"
              }
            },
            "Compose": {
              "runAfter": {
                "Analyze_Image": [
                  "Succeeded"
                ]
              },
              "metadata": {
                "operationMetadataId": "cbd09106-eb7a-4743-b553-8b11cd09a90c"
              },
              "type": "Compose",
              "inputs": "@outputs('Analyze_Image')?['body/tags']"
            },
            "Initialize_variable": {
              "runAfter": {},
              "metadata": {
                "operationMetadataId": "818f9da8-8ebf-4081-8d7b-8fe37470cd57"
              },
              "type": "InitializeVariable",
              "inputs": {
                "variables": [
                  {
                    "name": "FilePath",
                    "type": "string",
                    "value": "@{decodeUriComponent(decodeUriComponent(triggerOutputs()?['headers/x-ms-file-id']))}"
                  }
                ]
              }
            },
            "Send_an_HTTP_request_to_SharePoint": {
              "runAfter": {
                "Compose": [
                  "Succeeded"
                ]
              },
              "metadata": {
                "operationMetadataId": "0e9b2bf7-c52e-4b28-8351-2850a3373186"
              },
              "type": "OpenApiConnection",
              "inputs": {
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                  "connectionName": "shared_sharepointonline",
                  "operationId": "HttpRequest"
                },
                "parameters": {
                  "dataset": "https://contoso.sharepoint.com/sites/contosoHome",
                  "parameters/method": "GET",
                  "parameters/uri": "_api/web/getFileByServerRelativeURL('@{concat('/sites/contosoHome/', variables('FilePath'))}')?$select=ListItemAllFields/ID&$expand=ListItemAllFields"
                },
                "authentication": "@parameters('$authentication')"
              }
            },
            "Parse_JSON": {
              "runAfter": {
                "Send_an_HTTP_request_to_SharePoint": [
                  "Succeeded"
                ]
              },
              "metadata": {
                "operationMetadataId": "7d758c3f-f603-44d0-94b6-26cb08590fb3"
              },
              "type": "ParseJson",
              "inputs": {
                "content": "@body('Send_an_HTTP_request_to_SharePoint')",
                "schema": {
                  "type": "object",
                  "properties": {
                    "d": {
                      "type": "object",
                      "properties": {
                        "__metadata": {
                          "type": "object",
                          "properties": {
                            "id": {
                              "type": "string"
                            },
                            "uri": {
                              "type": "string"
                            },
                            "type": {
                              "type": "string"
                            }
                          }
                        },
                        "ListItemAllFields": {
                          "type": "object",
                          "properties": {
                            "__metadata": {
                              "type": "object",
                              "properties": {
                                "id": {
                                  "type": "string"
                                },
                                "uri": {
                                  "type": "string"
                                },
                                "etag": {
                                  "type": "string"
                                },
                                "type": {
                                  "type": "string"
                                }
                              }
                            },
                            "Id": {
                              "type": "integer"
                            },
                            "ID": {
                              "type": "integer"
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            },
            "Update_file_properties": {
              "runAfter": {
                "Parse_JSON": [
                  "Succeeded"
                ]
              },
              "metadata": {
                "operationMetadataId": "c1a801a8-1f96-498a-a356-82c9d0dd8188"
              },
              "type": "OpenApiConnection",
              "inputs": {
                "host": {
                  "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                  "connectionName": "shared_sharepointonline",
                  "operationId": "PatchFileItem"
                },
                "parameters": {
                  "dataset": "https://contoso.sharepoint.com/sites/contosoHome",
                  "table": "13e0b5a3-f626-4272-a73b-a3e84978199b",
                  "id": "@body('Parse_JSON')?['d']?['ListItemAllFields']?['Id']",
                  "item/CognitiveTags": "@join(outputs('Analyze_Image')?['body/description/tags'], ', ')"
                },
                "authentication": "@parameters('$authentication')"
              }
            }
          },
          "outputs": {}
        },
        "state": "Started",
        "connectionReferences": {
          "shared_cognitiveservicescomputervision": {
            "connectionName": "shared-cognitiveserv-34d89cec-971e-41b6-8c79-247c97eecec5",
            "source": "Embedded",
            "id": "/providers/Microsoft.PowerApps/apis/shared_cognitiveservicescomputervision",
            "displayName": "My Flow API",
            "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1549/1.0.1549.2680/cognitiveservicescomputervision/icon.png",
            "brandColor": "#1267AE",
            "tier": "Standard"
          },
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
        "createdTime": "2022-06-20T15:13:24.5135728Z",
        "lastModifiedTime": "2022-11-01T14:15:31.5790933Z",
        "flowSuspensionReason": "None",
        "environment": {
          "name": "Default-00000000-0000-0000-0000-000000000000",
          "type": "Microsoft.ProcessSimple/environments",
          "id": "/providers/Microsoft.ProcessSimple/environments/Default-00000000-0000-0000-0000-000000000000"
        },
        "definitionSummary": {
          "triggers": [
            {
              "type": "OpenApiConnection",
              "swaggerOperationId": "OnNewFile",
              "metadata": {
                "%252fComputer%2bVision%2bDemo": "/My Flow Demo",
                "operationMetadataId": "2e0aa276-5536-4f67-a7b7-c1a37a91b8ca",
                "%252fShared%2bDocuments": "/Shared Documents"
              }
            }
          ],
          "actions": [
            {
              "type": "OpenApiConnection",
              "swaggerOperationId": "AnalyzeImageV2",
              "metadata": {
                "operationMetadataId": "82b95160-9c96-4507-ad8c-c688ea95c2b1"
              }
            },
            {
              "type": "Compose",
              "metadata": {
                "operationMetadataId": "cbd09106-eb7a-4743-b553-8b11cd09a90c"
              }
            },
            {
              "type": "InitializeVariable",
              "metadata": {
                "operationMetadataId": "818f9da8-8ebf-4081-8d7b-8fe37470cd57"
              }
            },
            {
              "type": "OpenApiConnection",
              "swaggerOperationId": "HttpRequest",
              "metadata": {
                "operationMetadataId": "0e9b2bf7-c52e-4b28-8351-2850a3373186"
              }
            },
            {
              "type": "ParseJson",
              "metadata": {
                "operationMetadataId": "7d758c3f-f603-44d0-94b6-26cb08590fb3"
              }
            },
            {
              "type": "OpenApiConnection",
              "swaggerOperationId": "PatchFileItem",
              "metadata": {
                "operationMetadataId": "c1a801a8-1f96-498a-a356-82c9d0dd8188"
              }
            }
          ]
        },
        "creator": {
          "tenantId": "00000000-0000-0000-0000-000000000000",
          "objectId": "e1251b10-1ba4-49e3-b35a-933e3f21772b",
          "userId": "e1251b10-1ba4-49e3-b35a-933e3f21772b",
          "userType": "ActiveDirectory"
        },
        "installationStatus": "NotApplicable",
        "provisioningMethod": "FromDefinition",
        "flowFailureAlertSubscribed": true,
        "referencedResources": [
          {
            "service": "sharepoint",
            "resource": {
              "site": "https://contoso.sharepoint.com/sites/contosoHome"
            },
            "referencers": [
              {
                "referenceSourceType": "Triggers",
                "operationId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/apiOperations/OnNewFile"
              },
              {
                "referenceSourceType": "Actions",
                "operationId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/apiOperations/HttpRequest"
              }
            ]
          },
          {
            "service": "sharepoint",
            "resource": {
              "site": "https://contoso.sharepoint.com/sites/contosoHome",
              "list": "13e0b5a3-f626-4272-a73b-a3e84978199b"
            },
            "referencers": [
              {
                "referenceSourceType": "Actions",
                "operationId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/apiOperations/PatchFileItem"
              }
            ]
          }
        ],
        "isManaged": false
      },
      "displayName": "My Flow",
      "description": "",
      "triggers": "OpenApiConnection",
      "actions": "OpenApiConnection-AnalyzeImageV2, Compose, InitializeVariable, OpenApiConnection-HttpRequest, ParseJson, OpenApiConnection-PatchFileItem"
    }
    ```

=== "Text"

    ```text
    actions    : OpenApiConnection-AnalyzeImageV2, Compose, InitializeVariable, OpenApiConnection-HttpRequest, ParseJson, OpenApiConnection-PatchFileItem
    description:
    displayName: My Flow
    name       : ca76d7b8-3b76-4050-8c03-9fb310ad172f
    triggers   : OpenApiConnection
    ```

=== "CSV"

    ```csv
    name,displayName,description,triggers,actions
    ca76d7b8-3b76-4050-8c03-9fb310ad172f,My Flow,,OpenApiConnection,"OpenApiConnection-AnalyzeImageV2, Compose, InitializeVariable, OpenApiConnection-HttpRequest, ParseJson, OpenApiConnection-PatchFileItem"
    ```
