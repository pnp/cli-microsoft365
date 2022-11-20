# flow environment get

Gets information about the specified Microsoft Flow environment

## Usage

```sh
m365 flow environment get [options]
```

## Options

`-n, --name <name>`
: The name of the environment to get information about

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Get information about the Microsoft Flow environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_

```sh
m365 flow environment get --name Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

## Response

### Standard response

=== "JSON"

    ```json
    {
      "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
      "location": "india",
      "type": "Microsoft.ProcessSimple/environments",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
      "properties": {
        "displayName": "contoso (default)",
        "createdTime": "2019-12-21T18:32:11.8708704Z",
        "createdBy": {
          "id": "SYSTEM",
          "displayName": "SYSTEM",
          "type": "NotSpecified"
        },
        "provisioningState": "Succeeded",
        "creationType": "DefaultTenant",
        "environmentSku": "Default",
        "environmentType": "NotSpecified",
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
        "isDefault": true,
        "isPayAsYouGoEnabled": false,
        "azureRegionHint": "centralindia",
        "runtimeEndpoints": {
          "microsoft.BusinessAppPlatform": "https://india.api.bap.microsoft.com",
          "microsoft.CommonDataModel": "https://india.api.cds.microsoft.com",
          "microsoft.PowerApps": "https://india.api.powerapps.com",
          "microsoft.PowerAppsAdvisor": "https://india.api.advisor.powerapps.com",
          "microsoft.PowerVirtualAgents": "https://powervamg.in-il101.gateway.prod.island.powerapps.com",
          "microsoft.ApiManagement": "https://management.INDIA.azure-apihub.net",
          "microsoft.Flow": "https://india.api.flow.microsoft.com"
        },
        "linkedEnvironmentMetadata": {
          "type": "NotSpecified",
          "resourceId": "3aa550bf-52ac-42fc-98f7-5d1833c1501c",
          "friendlyName": "contoso (default)",
          "uniqueName": "orgfc80770f",
          "domainName": "orgfc80770f",
          "version": "9.2.22105.00154",
          "instanceUrl": "https://orgfc80770f.crm8.dynamics.com/",
          "instanceApiUrl": "https://orgfc80770f.api.crm8.dynamics.com",
          "baseLanguage": 1033,
          "instanceState": "Ready",
          "createdTime": "2019-12-25T15:46:14.433Z"
        },
        "environmentFeatures": {
          "isOpenApiEnabled": false
        },
        "cluster": {
          "category": "Prod",
          "number": "101",
          "uriSuffix": "in-il101.gateway.prod.island",
          "geoShortName": "IN",
          "environment": "Prod"
        },
        "governanceConfiguration": {
          "protectionLevel": "Basic"
        }
      },
      "displayName": "contoso (default)"
    }
    ```

=== "Text"

    ```text
    azureRegionHint  : centralindia
    displayName      : contoso (default)
    environmentSku   : Default
    id               : /providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5
    isDefault        : true
    location         : india
    name             : Default-d87a7535-dd31-4437-bfe1-95340acd55c5
    provisioningState: Succeeded
    ```

=== "CSV"

    ```csv
    name,id,location,displayName,provisioningState,environmentSku,azureRegionHint,isDefault
    Default-d87a7535-dd31-4437-bfe1-95340acd55c5,/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5,india,contoso (default),Succeeded,Default,centralindia,1  
    ```
