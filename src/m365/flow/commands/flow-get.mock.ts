export const flowGetResponse = {
  "name": "a18e89d1-4c75-41e4-9517-e90aedc079be",
  "id": "/providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/flows/a18e89d1-4c75-41e4-9517-e90aedc079be",
  "type": "Microsoft.ProcessSimple/environments/flows",
  "properties": {
    "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
    "displayName": "TEST",
    "userType": "Owner",
    "definition": {
      "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
      "contentVersion": "1.0.0.0",
      "parameters": {
        "$authentication": {
          "defaultValue": {},
          "type": "SecureObject"
        },
        "$connections": {
          "defaultValue": {},
          "type": "Object"
        }
      },
      "triggers": {
        "manual": {
          "metadata": {
            "operationMetadataId": "04077d6a-9fdd-4c79-b742-09a8d5c022eb"
          },
          "type": "Request",
          "kind": "Button",
          "inputs": {
            "schema": {
              "type": "object",
              "properties": {
                "text": {
                  "description": "Please enter your input",
                  "title": "Text",
                  "type": "string",
                  "x-ms-content-hint": "TEXT",
                  "x-ms-dynamically-added": true
                },
                "boolean": {
                  "description": "Please select yes or no",
                  "title": "Boolean",
                  "type": "boolean",
                  "x-ms-content-hint": "BOOLEAN",
                  "x-ms-dynamically-added": true
                },
                "number": {
                  "description": "Please enter a number",
                  "title": "Number",
                  "type": "number",
                  "x-ms-content-hint": "NUMBER",
                  "x-ms-dynamically-added": true
                }
              },
              "required": [
                "text",
                "boolean",
                "number"
              ]
            }
          }
        }
      },
      "actions": {
        "Initialize_variable": {
          "runAfter": {},
          "type": "InitializeVariable",
          "inputs": {
            "variables": [
              {
                "name": "TEST",
                "type": "string"
              }
            ]
          }
        }
      }
    },
    "triggerSchema": {
      "type": "object",
      "properties": {
        "text": {
          "description": "Please enter your input",
          "title": "Text",
          "type": "string",
          "x-ms-content-hint": "TEXT",
          "x-ms-dynamically-added": true
        },
        "boolean": {
          "description": "Please select yes or no",
          "title": "Boolean",
          "type": "boolean",
          "x-ms-content-hint": "BOOLEAN",
          "x-ms-dynamically-added": true
        },
        "number": {
          "description": "Please enter a number",
          "title": "Number",
          "type": "number",
          "x-ms-content-hint": "NUMBER",
          "x-ms-dynamically-added": true
        }
      },
      "required": [
        "text",
        "boolean",
        "number"
      ]
    },
    "state": "Started",
    "plan": "NotSpecified",
    "connectionReferences": {},
    "installedConnectionReferences": {
      "shared_approvals": {
        "connectionName": "shared-approvals-068230cf-43ef-455d-b823-ee6f989d0193",
        "apiDefinition": {
          "name": "shared_approvals",
          "id": "/providers/Microsoft.PowerApps/apis/shared_approvals",
          "type": "/providers/Microsoft.PowerApps/apis",
          "properties": {
            "displayName": "Approvals",
            "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1682/1.0.1682.3671/approvals/icon.png",
            "purpose": "NotSpecified",
            "connectionParameters": {},
            "runtimeUrls": [
              "https://europe-002.azure-apim.net/apim/approvals"
            ],
            "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/approvals",
            "metadata": {
              "source": "marketplace",
              "brandColor": "#6464F5",
              "connectionLimits": {
                "*": 1
              },
              "useNewApimVersion": "true",
              "version": {
                "previous": "releases/v1.0.1679\\1.0.1679.3643",
                "current": "releases/v1.0.1682\\1.0.1682.3671"
              }
            },
            "capabilities": [
              "actions"
            ],
            "tier": "Standard",
            "isCustomApi": false,
            "description": "Enables approvals in workflows.",
            "createdTime": "2018-09-16T08:09:23.9434372Z",
            "changedTime": "2024-03-11T18:25:46.7755101Z",
            "publisher": "Microsoft"
          }
        },
        "source": "Invoker",
        "id": "/providers/Microsoft.PowerApps/apis/shared_approvals",
        "displayName": "Approvals",
        "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1682/1.0.1682.3671/approvals/icon.png",
        "brandColor": "#6464F5",
        "tier": "Standard"
      }
    },
    "createdTime": "2023-07-25T18:42:15.7693633Z",
    "lastModifiedTime": "2024-03-08T10:24:12.0147046Z",
    "flowSuspensionReason": "None",
    "environment": {
      "name": "Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "type": "Microsoft.ProcessSimple/environments",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d"
    },
    "definitionSummary": {
      "triggers": [
        {
          "type": "OpenApiConnection",
          "swaggerOperationId": "GetOnUpdatedItems",
          "apiOperation": {
            "name": "GetOnUpdatedItems",
            "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/apiOperations/GetOnUpdatedItems",
            "type": "Microsoft.ProcessSimple/apis/apiOperations",
            "properties": {
              "summary": "When an item is created or modified",
              "description": "Triggers when an item is created, and also each time it is modified.",
              "visibility": "important",
              "trigger": "batch",
              "pageable": false,
              "isChunkingSupported": false,
              "isNotification": false,
              "annotation": {
                "status": "Production",
                "family": "GetOnUpdatedItems",
                "revision": 1
              },
              "externalDocs": {
                "url": "https://docs.microsoft.com/connectors/sharepointonline/#when-an-item-is-created-or-modified",
                "description": "Learn more"
              },
              "api": {
                "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
                "displayName": "SharePoint",
                "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1685/1.0.1685.3700/sharepointonline/icon.png",
                "brandColor": "#036C70",
                "tier": "Standard"
              },
              "operationType": "OpenApiConnection",
              "swaggerTags": [
                "SharePointListTableDataTrigger"
              ]
            }
          },
          "metadata": {
            "operationMetadataId": "0c830afe-5bfb-4af1-85bc-12ab218a1a2b"
          },
          "api": {
            "name": "shared_sharepointonline",
            "id": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
            "type": "/providers/Microsoft.PowerApps/apis",
            "properties": {
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1685/1.0.1685.3700/sharepointonline/icon.png",
              "metadata": {
                "source": "marketplace",
                "brandColor": "#036C70",
                "useNewApimVersion": "true",
                "version": {
                  "previous": "releases/v1.0.1682\\1.0.1682.3677",
                  "current": "releases/v1.0.1685\\1.0.1685.3700"
                }
              },
              "tier": "Standard",
              "isCustomApi": false,
              "description": "SharePoint helps organizations share and collaborate with colleagues, partners, and customers. You can connect to SharePoint Online or to an on-premises SharePoint 2016 or 2019 farm using the On-Premises Data Gateway to manage documents and list items."
            }
          }
        }
      ],
      "actions": [
        {
          "type": "InitializeVariable",
          "metadata": {
            "operationMetadataId": "629b46ae-5ede-4eb3-92b6-bc2ada4c51c0"
          }
        },
        {
          "type": "ApiConnection",
          "swaggerOperationId": "ListFeedItems",
          "metadata": {
            "flowSystemMetadata": {
              "swaggerOperationId": "ListFeedItems"
            }
          }
        }
      ]
    },
    "creator": {
      "tenantId": "e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "objectId": "fe36f75e-c103-410b-a18a-2bf6df06ac3a",
      "userId": "fe36f75e-c103-410b-a18a-2bf6df06ac3a",
      "userType": "ActiveDirectory"
    },
    "flowTriggerUri": "https://europe-002.azure-apim.net:443/apim/logicflows/A18E89D14C7541E49517E90AEDC079BE-5FA12B1209B98E3B/triggers/manual/run?api-version=2016-06-01",
    "installationStatus": "Installed",
    "provisioningMethod": "FromDefinition",
    "flowFailureAlertSubscribed": true,
    "referencedResources": [],
    "licenseData": {
      "performanceProfile": {
        "throttles": {
          "mode": "High"
        }
      },
      "flowLicenseName": "placeholder"
    },
    "isManaged": false,
    "machineDescriptionData": {},
    "flowOpenAiData": {
      "isConsequential": false,
      "isConsequentialFlagOverwritten": false
    }
  },
  "displayName": "TEST",
  "description": "",
  "triggers": "Request-Button",
  "actions": "InitializeVariable"
};