# pp environment get

Gets information about the specified Power Platform environment

## Usage

```sh
m365 pp environment get [options]
```

## Options

`-n, --name [name]`
: The name of the environment. When not specified, the default environment is retrieved.

`--asAdmin`
: Run the command as admin and retrieve details of environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.
    Register CLI for Microsoft 365 or Azure AD application as a management application for the Power Platform using 
    m365 pp managementapp add [options] 

## Examples

Get information about the Power Platform environment by name

```sh
m365 pp environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

Get information as admin about the Power Platform environment by name 

```sh
m365 pp environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
```

Get information about the default Power Platform environment

```sh
m365 pp environment get
```

## Response

=== "JSON"

    ```json
    {
      "id":"/providers/Microsoft.BusinessAppPlatform/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "type":"Microsoft.BusinessAppPlatform/environments",
      "location":"europe",
      "name":"Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "properties":{
        "tenantId":"e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
        "azureRegion":"northeurope",
        "displayName":"contoso (default)",
        "createdTime":"2020-03-12T13:39:17.9876946Z",
        "createdBy":{
          "id":"SYSTEM",
          "displayName":"SYSTEM",
          "type":"NotSpecified"
        },
        "provisioningState":"Succeeded",
        "creationType":"DefaultTenant",
        "environmentSku":"Default",
        "isDefault":true,
        "clientUris":{
          "admin":"https://admin.powerplatform.microsoft.com/environments/environment/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/hub",
          "maker":"https://make.powerapps.com/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/home"
        },
        "runtimeEndpoints":{
          "microsoft.BusinessAppPlatform":"https://europe.api.bap.microsoft.com",
          "microsoft.CommonDataModel":"https://europe.api.cds.microsoft.com",
          "microsoft.PowerApps":"https://europe.api.powerapps.com",
          "microsoft.PowerAppsAdvisor":"https://europe.api.advisor.powerapps.com",
          "microsoft.PowerVirtualAgents":"https://powervamg.eu-il109.gateway.prod.island.powerapps.com",
          "microsoft.ApiManagement":"https://management.EUROPE.azure-apihub.net",
          "microsoft.Flow":"https://emea.api.flow.microsoft.com"
        },
        "databaseType":"CommonDataService",
        "linkedEnvironmentMetadata":{
          "resourceId":"5041ef46-5a1c-4a0f-a185-6bb49b5c6686",
          "friendlyName":"contoso (default)",
          "uniqueName":"unq5041ef465a1c4a0fa1856bb49b5c6",
          "domainName":"org6633049a",
          "version":"9.2.22101.00168",
          "instanceUrl":"https://org6633049a.crm4.dynamics.com/",
          "instanceApiUrl":"https://org6633049a.api.crm4.dynamics.com",
          "baseLanguage":1033,
          "instanceState":"Ready",
          "createdTime":"2021-10-08T09:50:41.283Z",
          "backgroundOperationsState":"Enabled",
          "scaleGroup":"EURCRMLIVESG644",
          "platformSku":"Standard"
        },
        "trialScenarioType":"None",
        "retentionPeriod":"P7D",
        "states":{
          "management":{
            "id":"NotSpecified"
          },
          "runtime":{
            "runtimeReasonCode":"NotSpecified",
            "requestedBy":{
              "displayName":"SYSTEM",
              "type":"NotSpecified"
            },
            "id":"Enabled"
          }
        },
        "updateCadence":{
          "id":"Frequent"
        },
        "retentionDetails":{
          "retentionPeriod":"P7D",
          "backupsAvailableFromDateTime":"2022-10-23T23:33:19.6000451Z"
        },
        "protectionStatus":{
          "keyManagedBy":"Microsoft"
        },
        "cluster":{
          "category":"Prod",
          "number":"109",
          "uriSuffix":"eu-il109.gateway.prod.island",
          "geoShortName":"EU",
          "environment":"Prod"
        },
        "connectedGroups":[
          
        ],
        "lifecycleOperationsEnforcement":{
          "allowedOperations":[
            {
              "type":{
                "id":"Backup"
              }
            },
            {
              "type":{
                "id":"Edit"
              }
            },
            {
              "type":{
                "id":"Enable"
              }
            },
            {
              "type":{
                "id":"Disable"
              }
            },
            {
              "type":{
                "id":"EnableGovernanceConfiguration"
              }
            }
          ],
          "disallowedOperations":[
            {
              "type":{
                "id":"Provision"
              },
              "reason":{
                "message":"Provision cannot be performed because there is no linked CDS instance or the CDS instance version is not supported."
              }
            },
            {
              "type":{
                "id":"Unlock"
              },
              "reason":{
                "message":"Unlock cannot be performed because there is no linked CDS instance or the CDS instance version is not supported."
              }
            },
            {
              "type":{
                "id":"Convert"
              },
              "reason":{
                "message":"Convert cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Copy"
              },
              "reason":{
                "message":"Copy cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Delete"
              },
              "reason":{
                "message":"Delete cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Promote"
              },
              "reason":{
                "message":"Promote cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Recover"
              },
              "reason":{
                "message":"Recover cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Reset"
              },
              "reason":{
                "message":"Reset cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"Restore"
              },
              "reason":{
                "message":"Restore cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"UpdateProtectionStatus"
              },
              "reason":{
                "message":"UpdateProtectionStatus cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"NewCustomerManagedKey"
              },
              "reason":{
                "message":"NewCustomerManagedKey cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"RotateCustomerManagedKey"
              },
              "reason":{
                "message":"RotateCustomerManagedKey cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"RevertToMicrosoftKey"
              },
              "reason":{
                "message":"RevertToMicrosoftKey cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"NewNetworkInjection"
              },
              "reason":{
                "message":"NewNetworkInjection cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"SwapNetworkInjection"
              },
              "reason":{
                "message":"SwapNetworkInjection cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"RevertNetworkInjection"
              },
              "reason":{
                "message":"RevertNetworkInjection cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"NewIdentity"
              },
              "reason":{
                "message":"NewIdentity cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"SwapIdentity"
              },
              "reason":{
                "message":"SwapIdentity cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"RevertIdentity"
              },
              "reason":{
                "message":"RevertIdentity cannot be performed on environment of type Default."
              }
            },
            {
              "type":{
                "id":"DisableGovernanceConfiguration"
              },
              "reason":{
                "message":"DisableGovernanceConfiguration cannot be performed on Power Platform environment because of the governance configuration."
              }
            },
            {
              "type":{
                "id":"UpdateGovernanceConfiguration"
              },
              "reason":{
                "message":"UpdateGovernanceConfiguration cannot be performed on Power Platform environment because of the governance configuration."
              }
            }
          ]
        },
        "governanceConfiguration":{
          "protectionLevel":"Basic"
        }
      }
    }
    ```

=== "Text"

    ```text
    id  : /providers/Microsoft.BusinessAppPlatform/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d
    name: Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d
    ```

=== "CSV"

    ```csv
    name,id
    Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d,/providers/Microsoft.BusinessAppPlatform/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d
    ```
