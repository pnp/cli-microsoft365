# aad group get

Gets information about the specified Azure AD Group

## Usage

```sh
m365 aad group get [options]
```

## Options

`-i, --id [id]`
: The object Id of the Azure AD group. Specify either `id` or `title` but not both

`-t, --title [title]`
: The display name of the Azure AD group. Specify either `id` or `title` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get information about an Azure AD Group by id

```sh
m365 aad group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about an Azure AD Group by title

```sh
m365 aad group get --title "Finance"
```

## Response

=== "JSON"

    ```json
    {
      "id": "c541afac-508e-40c7-8880-5a601b41737b",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2022-11-13T19:16:32Z",
      "creationOptions": [
        "YammerProvisioning"
      ],
      "description": "This is the default group for everyone in the network",
      "displayName": "All Company",
      "expirationDateTime": null,
      "groupTypes": [
        "Unified"
      ],
      "isAssignableToRole": null,
      "mail": "allcompany@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "allcompany",
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
        "SPO:SPO_c3e7794d-0726-49ac-805b-2d6b0edaefdc@SPO_44744d00-3da0-45e5-9e28-da5ab48c61ac",
        "SMTP:allcompany@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2022-11-13T19:16:32Z",
      "resourceBehaviorOptions": [
        "CalendarMemberReadOnly"
      ],
      "resourceProvisioningOptions": [],
      "securityEnabled": false,
      "securityIdentifier": "S-1-12-1-4076547856-1079300050-1399127439-2879739702",
      "theme": null,
      "visibility": "Public",
      "onPremisesProvisioningErrors": []
    }
    ```

=== "Text"

    ```text
    classification               : null
    createdDateTime              : 2022-11-13T19:16:32Z
    creationOptions              : ["YammerProvisioning"]
    deletedDateTime              : null
    description                  : This is the default group for everyone in the network
    displayName                  : All Company
    expirationDateTime           : null
    groupTypes                   : ["Unified"]
    id                           : c541afac-508e-40c7-8880-5a601b41737b
    isAssignableToRole           : null
    mail                         : allcompany@contoso.onmicrosoft.com
    mailEnabled                  : true
    mailNickname                 : allcompany
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
    proxyAddresses               : ["SPO:SPO_c3e7794d-0726-49ac-805b-2d6b0edaefdc@SPO_44744d00-3da0-45e5-9e28-da5ab48c61ac","SMTP:allcompany@contoso.onmicrosoft.com"]
    renewedDateTime              : 2022-11-13T19:16:32Z
    resourceBehaviorOptions      : ["CalendarMemberReadOnly"]
    resourceProvisioningOptions  : []
    securityEnabled              : false
    securityIdentifier           : S-1-12-1-4076547856-1079300050-1399127439-2879739702
    theme                        : null
    visibility                   : Public
    ```

=== "CSV"

    ```csv
    id,deletedDateTime,classification,createdDateTime,creationOptions,description,displayName,expirationDateTime,groupTypes,isAssignableToRole,mail,mailEnabled,mailNickname,membershipRule,membershipRuleProcessingState,onPremisesDomainName,onPremisesLastSyncDateTime,onPremisesNetBiosName,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,preferredDataLocation,preferredLanguage,proxyAddresses,renewedDateTime,resourceBehaviorOptions,resourceProvisioningOptions,securityEnabled,securityIdentifier,theme,visibility,onPremisesProvisioningErrors
    c541afac-508e-40c7-8880-5a601b41737b,,,2022-11-13T19:16:32Z,"[""YammerProvisioning""]",This is the default group for everyone in the network,All Company,,"[""Unified""]",,allcompany@contoso.onmicrosoft.com,1,allcompany,,,,,,,,,,,"[""SPO:SPO_c3e7794d-0726-49ac-805b-2d6b0edaefdc@SPO_44744d00-3da0-45e5-9e28-da5ab48c61ac"",""SMTP:allcompany@contoso.onmicrosoft.com""]",2022-11-13T19:16:32Z,"[""CalendarMemberReadOnly""]",[],,S-1-12-1-4076547856-1079300050-1399127439-2879739702,,Public,[]
    ```
