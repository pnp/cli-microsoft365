# pa environment get

Gets information about the specified Power Apps environment

## Usage

```sh
m365 pa environment get [options]
```

## Options

`-n, --name [name]`
: The name of the environment. When not specified, the default environment is retrieved.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Get information about the default Power Apps environment

```sh
m365 pa environment get
```

Get information about the Power Apps environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
m365 pa environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

## Response

=== "JSON"

    ```json
    {
      "id": "/providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "name": "Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
      "location": "europe",
      "type": "Microsoft.PowerApps/environments",
      "properties": {
        "azureRegionHint": "westeurope",
        "displayName": "contoso (default)",
        "createdTime": "2020-03-12T13:39:17.9876946Z",
        "createdBy": {
          "id": "SYSTEM",
          "displayName": "SYSTEM",
          "type": "NotSpecified"
        },
        "provisioningState": "Succeeded",
        "creationType": "DefaultTenant",
        "environmentSku": "Default",
        "isDefault": true,
        "clientUris": {
          "admin": "https://admin.powerplatform.microsoft.com/environments/environment/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/hub",
          "maker": "https://make.powerapps.com/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d/home"
        },
        "runtimeEndpoints": {
          "microsoft.BusinessAppPlatform": "https://europe.api.bap.microsoft.com",
          "microsoft.CommonDataModel": "https://europe.api.cds.microsoft.com",
          "microsoft.PowerApps": "https://europe.api.powerapps.com",
          "microsoft.PowerAppsAdvisor": "https://europe.api.advisor.powerapps.com",
          "microsoft.PowerVirtualAgents": "https://powervamg.eu-il109.gateway.prod.island.powerapps.com",
          "microsoft.ApiManagement": "https://management.EUROPE.azure-apihub.net",
          "microsoft.Flow": "https://emea.api.flow.microsoft.com"
        },
        "databaseType": "CommonDataService",
        "linkedEnvironmentMetadata": {
          "resourceId": "5041ef46-5a1c-4a0f-a185-6bb49b5c6686",
          "friendlyName": "contoso (default)",
          "uniqueName": "unq5041ef465a1c4a0fa1856bb49b5c6",
          "domainName": "org6633049a",
          "version": "9.2.22101.00168",
          "instanceUrl": "https://org6633049a.crm4.dynamics.com/",
          "instanceApiUrl": "https://org6633049a.api.crm4.dynamics.com",
          "baseLanguage": 1033,
          "instanceState": "Ready",
          "createdTime": "2021-10-08T09:50:41.283Z",
          "platformSku": "Standard"
        },
        "retentionPeriod": "P7D",
        "lifecycleAuthority": "Environment",
        "states": {
          "management": {
            "id": "NotSpecified"
          },
          "runtime": {
            "runtimeReasonCode": "NotSpecified",
            "requestedBy": {
              "displayName": "SYSTEM",
              "type": "NotSpecified"
            },
            "id": "Enabled"
          }
        },
        "updateCadence": {
          "id": "Frequent"
        },
        "connectedGroups": [],
        "protectionStatus": {
          "keyManagedBy": "Microsoft"
        },
        "trialScenarioType": "None",
        "cluster": {
          "category": "Prod",
          "number": "109",
          "uriSuffix": "eu-il109.gateway.prod.island",
          "geoShortName": "EU",
          "environment": "Prod"
        },
        "governanceConfiguration": {
          "protectionLevel": "Basic"
        }
      },
      "displayName": "contoso (default)",
      "provisioningState": "Succeeded",
      "environmentSku": "Default",
      "azureRegionHint": "westeurope",
      "isDefault": true
    }
    ```

=== "Text"

    ```text
    azureRegionHint  : westeurope
    displayName      : environmentName (default)
    environmentSku   : Default
    id               : /providers/Microsoft.PowerApps/environments/Default-e1dd4030-a657-480a-8a0e-c1b1eec51e2e
    isDefault        : true
    location         : europe
    name             : Default-e1dd4030-a657-480a-8a0e-c1b1eec51e2e
    provisioningState: Succeeded
    ```

=== "CSV"

    ```csv
    name,id,location,displayName,provisioningState,environmentSku,azureRegionHint,isDefault
    Default-e1dd4030-a657-480a-8a0e-c1b1eec51e2e,/providers/Microsoft.PowerApps/environments/Default-e1dd4030-a657-480a-8a0e-c1b1eec51e2e,europe,environmentName (default),Succeeded,Default,westeurope,1
    ```
    
=== "Markdown"

    ```md
    # pa environment get --name "Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d"

    Date: 9/1/2023

    ## environmentName (default) (org6633050c) (/providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d)

    Property | Value
    ---------|-------
    id | /providers/Microsoft.PowerApps/environments/Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d
    name | Default-e1dd4023-a656-480a-8a0e-c1b1eec51e1d
    location | europe
    type | Microsoft.PowerApps/environments
    properties | {"azureRegionHint":"westeurope","displayName":"contoso (default) (org6633049a)","createdTime":"2020-03-12T13:39:17.9876946Z","createdBy":{"id":"SYSTEM","displayName":"SYSTEM","type":"NotSpecified"},"provisioningState":"Succeeded","creationType":"DefaultTenant","environmentSku":"Default","environmentType":"Production","isDefault":true,"runtimeEndpoints":{"microsoft.BusinessAppPlatform":"https://europe.api.bap.microsoft.com","microsoft.CommonDataModel":"https://europe.api.cds.microsoft.com","microsoft.PowerApps":"https://europe.api.powerapps.com","microsoft.PowerAppsAdvisor":"https://europe.api.advisor.powerapps.com","microsoft.PowerVirtualAgents":"https://powervamg.eu-il109.gateway.prod.island.powerapps.com","microsoft.ApiManagement":"https://management.EUROPE.azure-apihub.net","microsoft.Flow":"https://emea.api.flow.microsoft.com"},"linkedEnvironmentMetadata":{"type":"Dynamics365Instance","resourceId":"5041ef46-5a1c-4a0f-a185-6bb49b5c6686","friendlyName":"contoso (default)","uniqueName":"unq5041ef465a1c4a0fa1856bb49b5c6","domainName":"org6633049a","version":"9.2.22101.00168","instanceUrl":"https://org6633049a.crm4.dynamics.com/","instanceApiUrl":"https://org6633049a.api.crm4.dynamics.com","baseLanguage":1033,"instanceState":"Ready","createdTime":"2021-10-08T09:50:41.283Z","modifiedTime":"2022-10-29T14:04:14.0720726Z","hostNameSuffix":"crm4.dynamics.com","bapSolutionId":"00000001-0000-0000-0001-00000000009b","creationTemplates":["D365_CDS"],"webApiVersion":"v9.0","platformSku":"Standard"},"retentionPeriod":"P7D","lifecycleAuthority":"Environment","states":{"management":{"id":"NotSpecified"},"runtime":{"runtimeReasonCode":"NotSpecified","requestedBy":{"displayName":"SYSTEM","type":"NotSpecified"},"id":"Enabled"}},"updateCadence":{"id":"Frequent"},"connectedGroups":[],"protectionStatus":{"keyManagedBy":"Microsoft"},"trialScenarioType":"None","cluster":{"category":"Prod","number":"109","uriSuffix":"eu-il109.gateway.prod.island","geoShortName":"EU","environment":"Prod"},"governanceConfiguration":{"protectionLevel":"Basic"}}
    displayName | contoso (default) (org6633049a)
    ```
