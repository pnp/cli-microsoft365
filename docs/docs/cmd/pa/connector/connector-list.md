# pa connector list

Lists custom connectors in the given environment

## Usage

```sh
m365 pa connector list [options]
```

## Alias

```sh
m365 flow connector list
```

## Options

`-e, --environmentName <environmentName>`
: The name of the environment for which to retrieve custom connectors

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

List all custom connectors in the given environment

```sh
m365 pa connector list --environmentName Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```

## Response

=== "JSON"

    ```json
    [
      {
        "name":"shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012",
        "id":"/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012",
        "type":"Microsoft.PowerApps/apis",
        "properties":{
          "displayName":"My connector",
          "iconUri":"https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png",
          "iconBrandColor":"#007ee5",
          "contact":{},
          "license":{},
          "apiEnvironment":"Shared",
          "isCustomApi":true,
          "connectionParameters":{},
          "runtimeUrls":[
            "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"
          ],
          "primaryRuntimeUrl":"https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012",
          "metadata":{
            "source":"powerapps-user-defined",
            "brandColor":"#007ee5",
            "contact":{},
            "license":{},
            "publisherUrl":null,
            "serviceUrl":null,
            "documentationUrl":null,
            "environmentName":"Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6",
            "xrmConnectorId":null,
            "almMode":"Environment",
            "createdBy":"{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
            "modifiedBy":"{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
            "allowSharing":false
          },
          "capabilities":[],
          "description":"",
          "apiDefinitions":{
            "originalSwaggerUrl":"https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r",
            "modifiedSwaggerUrl":"https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r"
          },
          "createdBy":{
            "id":"03043611-d01e-4e58-9fbe-1a18ecb861d8",
            "displayName":"MOD Administrator",
            "email":"admin@contoso.OnMicrosoft.com",
            "type":"User",
            "tenantId":"0d645e38-ec52-4a4f-ac58-65f2ac4015f6",
            "userPrincipalName":"admin@contoso.onmicrosoft.com"
          },
          "modifiedBy":{
            "id":"03043611-d01e-4e58-9fbe-1a18ecb861d8",
            "displayName":"MOD Administrator",
            "email":"admin@contoso.OnMicrosoft.com",
            "type":"User",
            "tenantId":"0d645e38-ec52-4a4f-ac58-65f2ac4015f6",
            "userPrincipalName":"admin@contoso.onmicrosoft.com"
          },
          "createdTime":"2019-12-05T18:45:03.4615313Z",
          "changedTime":"2019-12-05T18:45:03.4615313Z",
          "environment":{
            "id":"/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6",
            "name":"Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6"
          },
          "tier":"Standard",
          "publisher":"MOD Administrator",
          "almMode":"Environment"
        }
      }
    ]
    ```

=== "Text"

    ```text
    name                                                        displayName
    ----------------------------------------------------------- ------------
    shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012 My connector
    ```

=== "CSV"

    ```csv
    name,displayName
    shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012,My connector
    ```
