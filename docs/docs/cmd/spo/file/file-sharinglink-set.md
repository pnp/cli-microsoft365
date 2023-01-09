# spo file sharinglink set

Updates a sharing link from a file

## Usage

```sh
m365 spo file sharinglink set [options]
```

## Options

`-w, --webUrl <webUrl>`
:	The URL of the site where the file is located.

`-u, --fileUrl [fileUrl]`
:	The server-relative URL of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --fileId [fileId]`
:	The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`--id <id>`
: The ID of the sharing link.

`--role [role]`
: The role to set. Possible options are: `read` or `write`. Specify either `role` or `expirationDateTime` but not both.

`--expirationDateTime [expirationDateTime]`
:	The date and time to set the expiration. This should be defined as a valid ISO 8601 string. Specify either `role` or `expirationDateTime` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks
- you can only set the expiration date for anonymous sharing links
- you can only set the role for user sharing

## Examples

Updates a user sharing link from a file with the role parameter

```sh
m365 spo file sharinglink set --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --role read
```

Updates an anonymous sharing link from a file with the expirationDateTime parameter

```sh
m365 spo file sharinglink set --webUrl https://contoso.sharepoint.com/sites/demo --fileUrl "/Shared Documents/Document.docx" --expirationDateTime "2023-01-09"
```

## Response

=== "JSON"

    ```json
    {
      "__metadata": {
        "type": "SP.SharingLinkInfo"
      },
      "AllowsAnonymousAccess": true,
      "ApplicationId": null,
      "BlocksDownload": false,
      "Created": "2023-01-09T20:20:22.999Z",
      "CreatedBy": {
        "__metadata": {
          "type": "SP.Sharing.Principal"
        },
        "directoryObjectId": null,
        "email": "john@contoso.onmicrosoft.com",
        "expiration": null,
        "id": 10,
        "isActive": true,
        "isExternal": false,
        "jobTitle": null,
        "loginName": "i:0#.f|membership|john@contoso.onmicrosoft.com",
        "name": "John Doe",
        "principalType": 1,
        "userId": null,
        "userPrincipalName": "john@contoso.onmicrosoft.com"
      },
      "Description": null,
      "Embeddable": false,
      "Expiration": "2023-10-31T23:00:00.000Z",
      "HasExternalGuestInvitees": false,
      "Invitations": {
        "__metadata": {
          "type": "Collection(SP.Sharing.LinkInvitation)"
        },
        "results": []
      },
      "IsActive": true,
      "IsAddressBarLink": false,
      "IsCreateOnlyLink": false,
      "IsDefault": true,
      "IsEditLink": false,
      "IsFormsLink": false,
      "IsManageListLink": false,
      "IsReviewLink": false,
      "IsUnhealthy": false,
      "LastModified": "2023-01-09T21:14:53.181Z",
      "LastModifiedBy": {
        "__metadata": {
          "type": "SP.Sharing.Principal"
        },
        "directoryObjectId": null,
        "email": "john@contoso.onmicrosoft.com",
        "expiration": null,
        "id": 10,
        "isActive": true,
        "isExternal": false,
        "jobTitle": null,
        "loginName": "i:0#.f|membership|john@contoso.onmicrosoft.com",
        "name": "John Doe",
        "principalType": 1,
        "userId": null,
        "userPrincipalName": "john@contoso.onmicrosoft.com"
      },
      "LimitUseToApplication": false,
      "LinkKind": 4,
      "PasswordLastModified": "",
      "PasswordLastModifiedBy": null,
      "RedeemedUsers": {
        "__metadata": {
          "type": "Collection(SP.Sharing.LinkInvitation)"
        },
        "results": []
      },
      "RequiresPassword": false,
      "RestrictedShareMembership": false,
      "Scope": 0,
      "ShareId": "7c9f97c9-1bda-433c-9364-bb83e81771ee",
      "ShareTokenString": "share=EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g",
      "SharingLinkStatus": 2,
      "TrackLinkUsers": false,
      "Url": "https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g"
    }
    ```

=== "Text"

    ```text
    id   : 7c9f97c9-1bda-433c-9364-bb83e81771ee
    link : https://ordidev.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g
    scope: 0
    ```

=== "CSV"

    ```csv
    id,link,scope
    7c9f97c9-1bda-433c-9364-bb83e81771ee,https://ordidev.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g,0
    ```

=== "Markdown"

    ```md
    # spo file sharinglink set --webUrl "https://contoso.sharepoint.com" --fileUrl "/Shared Documents/Document.docx" --id "7c9f97c9-1bda-433c-9364-bb83e81771ee" --expirationDateTime "2023-11-01"

    Date: 9/1/2023

    ## undefined (https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g)

    Property | Value
    ---------|-------
    \_\_metadata | {"type":"SP.SharingLinkInfo"}
    AllowsAnonymousAccess | true
    ApplicationId | null
    BlocksDownload | false
    Created | 2023-01-09T20:20:22.999Z
    CreatedBy | {"\_\_metadata":{"type":"SP.Sharing.Principal"},"directoryObjectId":null,"email":"john@contoso.onmicrosoft.com","expiration":null,"id":10,"isActive":true,"isExternal":false,"jobTitle":null,"loginName":"i:0#.f\|membership\|john@contoso.onmicrosoft.com","name":"John Doe","principalType":1,"userId":null,"userPrincipalName":"john@contoso.onmicrosoft.com"}
    Description | null
    Embeddable | false
    Expiration | 2023-10-31T23:00:00.000Z
    HasExternalGuestInvitees | false
    Invitations | {"\_\_metadata":{"type":"Collection(SP.Sharing.LinkInvitation)"},"results":[]}
    IsActive | true
    IsAddressBarLink | false
    IsCreateOnlyLink | false
    IsDefault | true
    IsEditLink | false
    IsFormsLink | false
    IsManageListLink | false
    IsReviewLink | false
    IsUnhealthy | false
    LastModified | 2023-01-09T21:14:53.181Z
    LastModifiedBy | {"\_\_metadata":{"type":"SP.Sharing.Principal"},"directoryObjectId":null,"email":"john@contoso.onmicrosoft.com","expiration":null,"id":10,"isActive":true,"isExternal":false,"jobTitle":null,"loginName":"i:0#.f\|membership\|john@contoso.onmicrosoft.com","name":"John Doe","principalType":1,"userId":null,"userPrincipalName":"john@contoso.onmicrosoft.com"}
    LimitUseToApplication | false
    LinkKind | 4
    PasswordLastModified |
    PasswordLastModifiedBy | null
    RedeemedUsers | {"\_\_metadata":{"type":"Collection(SP.Sharing.LinkInvitation)"},"results":[]}
    RequiresPassword | false
    RestrictedShareMembership | false
    Scope | 0
    ShareId | 7c9f97c9-1bda-433c-9364-bb83e81771ee
    ShareTokenString | share=EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-\_527\_fI9g
    SharingLinkStatus | 2
    TrackLinkUsers | false
    Url | https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-\_527\_fI9g
    ```
