# app get

Retrieves information about the current Azure AD app

## Usage

```sh
m365 app get [options]
```

## Options

`--appId [appId]`
: Client ID of the Azure AD app registered in the .m365rc.json file to retrieve information for

--8<-- "docs/cmd/_global.md"

## Remarks

Use this command to quickly look up information for the Azure AD application registration registered in the .m365rc.json file in your current project (folder).

If you have multiple apps registered in your .m365rc.json file, you can specify the app for which you'd like to retrieve permissions using the `--appId` option. If you don't specify the app using the `--appId` option, you'll be prompted to select one of the applications from your .m365rc.json file.

## Examples

Retrieve information about your current Azure AD app

```sh
m365 app get
```

Retrieve information about the Azure AD app with client ID _e23d235c-fcdf-45d1-ac5f-24ab2ee0695d_ specified in the _.m365rc.json_ file

```sh
m365 app get --appId e23d235c-fcdf-45d1-ac5f-24ab2ee0695d
```

## Response

=== "JSON"

    ```json
    {
      "id": "b3241f2a-6551-43d4-a72f-0d02ccfeb77f",
      "deletedDateTime": null,
      "appId": "02f9ff8c-b79e-4552-bdda-4facd74d6df1",
      "applicationTemplateId": null,
      "disabledByMicrosoftStatus": null,
      "createdDateTime": "2022-10-28T21:20:57Z",
      "displayName": "My AAD app",
      "description": null,
      "groupMembershipClaims": null,
      "identifierUris": [],
      "isDeviceOnlyAuthSupported": null,
      "isFallbackPublicClient": null,
      "notes": null,
      "publisherDomain": "8b7jz1.onmicrosoft.com",
      "serviceManagementReference": null,
      "signInAudience": "AzureADMyOrg",
      "tags": [],
      "tokenEncryptionKeyId": null,
      "samlMetadataUrl": null,
      "defaultRedirectUri": null,
      "certification": null,
      "optionalClaims": null,
      "addIns": [],
      "api": {
        "acceptMappedClaims": null,
        "knownClientApplications": [],
        "requestedAccessTokenVersion": null,
        "oauth2PermissionScopes": [],
        "preAuthorizedApplications": []
      },
      "appRoles": [
        {
          "allowedMemberTypes": [
            "User"
          ],
          "description": "Managers",
          "displayName": "Managers",
          "id": "d7c46856-8a66-485a-9047-098eba08b9a3",
          "isEnabled": true,
          "origin": "Application",
          "value": "managers"
        }
      ],
      "info": {
        "logoUrl": null,
        "marketingUrl": null,
        "privacyStatementUrl": null,
        "supportUrl": null,
        "termsOfServiceUrl": null
      },
      "keyCredentials": [],
      "parentalControlSettings": {
        "countriesBlockedForMinors": [],
        "legalAgeGroupRule": "Allow"
      },
      "passwordCredentials": [],
      "publicClient": {
        "redirectUris": []
      },
      "requiredResourceAccess": [
        {
          "resourceAppId": "b55b276d-2b09-4ad2-8de5-f09cf24ffba9",
          "resourceAccess": [
            {
              "id": "162b1576-a2b2-458d-b7b9-04481911b4ef",
              "type": "Role"
            }
          ]
        },
        {
          "resourceAppId": "00000005-0000-0ff1-ce00-000000000000",
          "resourceAccess": [
            {
              "id": "8e5870bb-8808-44dc-8e10-c509ed919ddd",
              "type": "Scope"
            },
            {
              "id": "5db81a03-0de0-432b-b31e-71d57c8d2e0b",
              "type": "Scope"
            }
          ]
        }
      ],
      "verifiedPublisher": {
        "displayName": null,
        "verifiedPublisherId": null,
        "addedDateTime": null
      },
      "web": {
        "homePageUrl": null,
        "logoutUrl": null,
        "redirectUris": [],
        "implicitGrantSettings": {
          "enableAccessTokenIssuance": false,
          "enableIdTokenIssuance": false
        },
        "redirectUriSettings": []
      },
      "spa": {
        "redirectUris": []
      }
    }
    ```

=== "Text"

    ```text
    addIns                    : []
    api                       : {"acceptMappedClaims":null,"knownClientApplications":[],"requestedAccessTokenVersion":null,"oauth2PermissionScopes":[],"preAuthorizedApplications":[]}
    appId                     : 02f9ff8c-b79e-4552-bdda-4facd74d6df1
    appRoles                  : [{"allowedMemberTypes":["User"],"description":"Managers","displayName":"Managers","id":"d7c46856-8a66-485a-9047-098eba08b9a3","isEnabled":true,"origin":"Application","value":"managers"}]
    applicationTemplateId     : null
    certification             : null
    createdDateTime           : 2022-10-28T21:20:57Z
    defaultRedirectUri        : null
    deletedDateTime           : null
    description               : null
    disabledByMicrosoftStatus : null
    displayName               : My AAD app
    groupMembershipClaims     : null
    id                        : b3241f2a-6551-43d4-a72f-0d02ccfeb77f
    identifierUris            : []
    info                      : {"logoUrl":null,"marketingUrl":null,"privacyStatementUrl":null,"supportUrl":null,"termsOfServiceUrl":null}
    isDeviceOnlyAuthSupported : null
    isFallbackPublicClient    : null
    keyCredentials            : []
    notes                     : null
    optionalClaims            : null
    parentalControlSettings   : {"countriesBlockedForMinors":[],"legalAgeGroupRule":"Allow"}
    passwordCredentials       : []
    publicClient              : {"redirectUris":[]}
    publisherDomain           : 8b7jz1.onmicrosoft.com
    requiredResourceAccess    : [{"resourceAppId":"b55b276d-2b09-4ad2-8de5-f09cf24ffba9","resourceAccess":[{"id":"162b1576-a2b2-458d-b7b9-04481911b4ef","type":"Role"}]},{"resourceAppId":"00000005-0000-0ff1-ce00-000000000000","resourceAccess":[{"id":"8e5870bb-8808-44dc-8e10-c509ed919ddd","type":"Scope"},{"id":"5db81a03-0de0-432b-b31e-71d57c8d2e0b","type":"Scope"}]}]
    samlMetadataUrl           : null
    serviceManagementReference: null
    signInAudience            : AzureADMyOrg
    spa                       : {"redirectUris":[]}
    tags                      : []
    tokenEncryptionKeyId      : null
    verifiedPublisher         : {"displayName":null,"verifiedPublisherId":null,"addedDateTime":null}
    web                       : {"homePageUrl":null,"logoutUrl":null,"redirectUris":[],"implicitGrantSettings":{"enableAccessTokenIssuance":false,"enableIdTokenIssuance":false},"redirectUriSettings":[]}
    ```

=== "CSV"

    ```csv
    id,deletedDateTime,appId,applicationTemplateId,disabledByMicrosoftStatus,createdDateTime,displayName,description,groupMembershipClaims,identifierUris,isDeviceOnlyAuthSupported,isFallbackPublicClient,notes,publisherDomain,serviceManagementReference,signInAudience,tags,tokenEncryptionKeyId,samlMetadataUrl,defaultRedirectUri,certification,optionalClaims,addIns,api,appRoles,info,keyCredentials,parentalControlSettings,passwordCredentials,publicClient,requiredResourceAccess,verifiedPublisher,web,spa
    b3241f2a-6551-43d4-a72f-0d02ccfeb77f,,02f9ff8c-b79e-4552-bdda-4facd74d6df1,,,2022-10-28T21:20:57Z,My AAD app,,,[],,,,8b7jz1.onmicrosoft.com,,AzureADMyOrg,[],,,,,,[],"{""acceptMappedClaims"":null,""knownClientApplications"":[],""requestedAccessTokenVersion"":null,""oauth2PermissionScopes"":[],""preAuthorizedApplications"":[]}","[{""allowedMemberTypes"":[""User""],""description"":""Managers"",""displayName"":""Managers"",""id"":""d7c46856-8a66-485a-9047-098eba08b9a3"",""isEnabled"":true,""origin"":""Application"",""value"":""managers""}]","{""logoUrl"":null,""marketingUrl"":null,""privacyStatementUrl"":null,""supportUrl"":null,""termsOfServiceUrl"":null}",[],"{""countriesBlockedForMinors"":[],""legalAgeGroupRule"":""Allow""}",[],"{""redirectUris"":[]}","[{""resourceAppId"":""b55b276d-2b09-4ad2-8de5-f09cf24ffba9"",""resourceAccess"":[{""id"":""162b1576-a2b2-458d-b7b9-04481911b4ef"",""type"":""Role""}]},{""resourceAppId"":""00000005-0000-0ff1-ce00-000000000000"",""resourceAccess"":[{""id"":""8e5870bb-8808-44dc-8e10-c509ed919ddd"",""type"":""Scope""},{""id"":""5db81a03-0de0-432b-b31e-71d57c8d2e0b"",""type"":""Scope""}]}]","{""displayName"":null,""verifiedPublisherId"":null,""addedDateTime"":null}","{""homePageUrl"":null,""logoutUrl"":null,""redirectUris"":[],""implicitGrantSettings"":{""enableAccessTokenIssuance"":false,""enableIdTokenIssuance"":false},""redirectUriSettings"":[]}","{""redirectUris"":[]}"
    ```
