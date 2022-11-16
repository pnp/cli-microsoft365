# spo list label get

Gets label set on the specified list

## Usage

```sh
m365 spo list label get  [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to get the label from is located

`-l, --listId [listId]`
: ID of the list to get the label from. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to get the label from. Specify either `listId` or `listTitle` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Gets label set on the list with title _ContosoList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label get  --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle ContosoList
```

Gets label set on the list with id _cc27a922-8224-4296-90a5-ebbc54da2e85_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label get  --webUrl https://contoso.sharepoint.com/sites/project-x --listId cc27a922-8224-4296-90a5-ebbc54da2e85
```

## Response

=== "JSON"

    ```json
    {
      "AcceptMessagesOnlyFromSendersOrMembers": false,
      "AccessType": null,
      "AllowAccessFromUnmanagedDevice": null,
      "AutoDelete": false,
      "BlockDelete": false,
      "BlockEdit": false,
      "ContainsSiteLabel": false,
      "DisplayName": "Label A",
      "EncryptionRMSTemplateId": null,
      "HasRetentionAction": false,
      "IsEventTag": false,
      "Notes": null,
      "RequireSenderAuthenticationEnabled": false,
      "ReviewerEmail": null,
      "SharingCapabilities": null,
      "SuperLock": false,
      "TagDuration": 0,
      "TagId": "4d535433-2a7b-40b0-9dad-8f0f8f3b3841",
      "TagName": "Sensitive",
      "TagRetentionBasedOn": null
    }
    ```

=== "Text"

    ```text
    AcceptMessagesOnlyFromSendersOrMembers: false
    AccessType                            : null
    AllowAccessFromUnmanagedDevice        : null
    AutoDelete                            : false
    BlockDelete                           : false
    BlockEdit                             : false
    ContainsSiteLabel                     : false
    DisplayName                           : Label A
    EncryptionRMSTemplateId               : null
    HasRetentionAction                    : false
    IsEventTag                            : false
    Notes                                 : null
    RequireSenderAuthenticationEnabled    : false
    ReviewerEmail                         : null
    SharingCapabilities                   : null
    SuperLock                             : false
    TagDuration                           : 0
    TagId                                 : 4d535433-2a7b-40b0-9dad-8f0f8f3b3841
    TagName                               : Sensitive
    TagRetentionBasedOn                   : null
    ```

=== "CSV"

    ```csv
    AcceptMessagesOnlyFromSendersOrMembers,AccessType,AllowAccessFromUnmanagedDevice,AutoDelete,BlockDelete,BlockEdit,ContainsSiteLabel,DisplayName,EncryptionRMSTemplateId,HasRetentionAction,IsEventTag,Notes,RequireSenderAuthenticationEnabled,ReviewerEmail,SharingCapabilities,SuperLock,TagDuration,TagId,TagName,TagRetentionBasedOn
    false,,,false,false,false,false,Label A,,false,false,,false,,,false,0,4d535433-2a7b-40b0-9dad-8f0f8f3b3841,Sensitive,
    ```
