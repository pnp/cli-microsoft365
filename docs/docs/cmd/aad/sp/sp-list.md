# aad sp list

Lists the service principals in the directory

## Usage

```sh
m365 aad sp list [options]
```

## Options

`--displayName [displayName]`
: Returns only service principals with the specified name

`--tags [tag]`
:	Returns only service principals with the specified tag

--8<-- "docs/cmd/_global.md"

## Examples

Return a list of all service principals

```sh
m365 aad sp list
```

Return a list of all service principals that comply to the displayName and the tags parameters

```sh
m365 aad sp list --displayName "My custom service principal" --tags "WindowsAzureActiveDirectoryIntegratedApp,disableRequestingTenantedPassthroughTokens"
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "226859cc-86f0-40d3-b308-f43b3a729b6e",
        "deletedDateTime": null,
        "accountEnabled": true,
        "alternativeNames": [],
        "appDisplayName": "My custom service principal",
        "appDescription": null,
        "appId": "a62ef842-f9ef-49cf-9119-31b85ea58445",
        "applicationTemplateId": null,
        "appOwnerOrganizationId": "fd71909b-55e5-44d2-9f78-dc432421d527",
        "appRoleAssignmentRequired": false,
        "createdDateTime": "2022-11-28T20:32:11Z",
        "description": null,
        "disabledByMicrosoftStatus": null,
        "displayName": "My custom service principal",
        "homepage": null,
        "loginUrl": null,
        "logoutUrl": null,
        "notes": null,
        "notificationEmailAddresses": [],
        "preferredSingleSignOnMode": null,
        "preferredTokenSigningKeyThumbprint": null,
        "replyUrls": [
          "urn:ietf:wg:oauth:2.0:oob",
          "https://localhost",
          "http://localhost",
          "http://localhost:8400"
        ],
        "servicePrincipalNames": [
          "https://contoso.onmicrosoft.com/907a8cea-411a-461a-bb30-261e52febcca",
          "907a8cea-411a-461a-bb30-261e52febcca"
        ],
        "servicePrincipalType": "Application",
        "signInAudience": "AzureADMultipleOrgs",
        "tags": [
          "WindowsAzureActiveDirectoryIntegratedApp"
        ],
        "tokenEncryptionKeyId": null,
        "samlSingleSignOnSettings": null,
        "addIns": [],
        "appRoles": [],
        "info": {
          "logoUrl": null,
          "marketingUrl": null,
          "privacyStatementUrl": null,
          "supportUrl": null,
          "termsOfServiceUrl": null
        },
        "keyCredentials": [],
        "oauth2PermissionScopes": [
          {
            "adminConsentDescription": "Allow the application to access My custom service principal on behalf of the signed-in user.",
            "adminConsentDisplayName": "Access My custom service principal",
            "id": "907a8cea-411a-461a-bb30-261e52febcca",
            "isEnabled": true,
            "type": "User",
            "userConsentDescription": "Allow the application to access My custom service principal on your behalf.",
            "userConsentDisplayName": "Access My custom service principal",
            "value": "user_impersonation"
          }
        ],
        "passwordCredentials": [],
        "resourceSpecificApplicationPermissions": [],
        "verifiedPublisher": {
          "displayName": null,
          "verifiedPublisherId": null,
          "addedDateTime": null
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                   displayName                   tags
    --------------------------------------  ----------------------------  ---------------------------------------
    a62ef842-f9ef-49cf-9119-31b85ea58445    My custom service principal   WindowsAzureActiveDirectoryIntegratedApp
    ```

=== "CSV"

    ```csv
    id,accountEnabled,appDisplayName,appId,appOwnerOrganizationId,appRoleAssignmentRequired,createdDateTime,displayName,servicePrincipalType,signInAudience
    226859cc-86f0-40d3-b308-f43b3a729b6e,1,My custom service principal,a62ef842-f9ef-49cf-9119-31b85ea58445,fd71909b-55e5-44d2-9f78-dc432421d527,,2022-11-28T20:32:11Z,My custom service principal,AzureADMultipleOrgs
    ```
    
=== "Markdown"

    ```md
    # aad sp list

    Date: 27/4/2023

    ## My custom service principal (226859cc-86f0-40d3-b308-f43b3a729b6e)

    Property | Value
    ---------|-------
    id | 226859cc-86f0-40d3-b308-f43b3a729b6e
    accountEnabled | true
    appDisplayName | My custom service principal
    appId | a62ef842-f9ef-49cf-9119-31b85ea58445
    appOwnerOrganizationId | fd71909b-55e5-44d2-9f78-dc432421d527
    appRoleAssignmentRequired | false
    createdDateTime | 2022-11-28T20:32:11Z
    displayName | My custom service principal
    servicePrincipalType | Application
    signInAudience | AzureADMultipleOrg
    ```
