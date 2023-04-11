# spo web retentionlabel list

Get a list of retention labels that are available on a site.

## Usage

```sh
m365 spo web retentionlabel list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site.

--8<-- "docs/cmd/_global.md"

## Examples

Get a list of retention labels for the _Sales_ site

```sh
m365 spo web retentionlabel list --webUrl 'https://contoso.sharepoint.com/sites/sales'
```

## Response

=== "JSON"

    ```json
    [
      {
        "AcceptMessagesOnlyFromSendersOrMembers": false,
        "AccessType": null,
        "AllowAccessFromUnmanagedDevice": null,
        "AutoDelete": true,
        "BlockDelete": true,
        "BlockEdit": false,
        "ComplianceFlags": 1,
        "ContainsSiteLabel": false,
        "DisplayName": "",
        "EncryptionRMSTemplateId": null,
        "HasRetentionAction": true,
        "IsEventTag": false,
        "MultiStageReviewerEmail": null,
        "NextStageComplianceTag": null,
        "Notes": null,
        "RequireSenderAuthenticationEnabled": false,
        "ReviewerEmail": null,
        "SharingCapabilities": null,
        "SuperLock": false,
        "TagDuration": 2555,
        "TagId": "def61080-111c-4aea-b72f-5b60e516e36c",
        "TagName": "Some label",
        "TagRetentionBasedOn": "CreationAgeInDays",
        "UnlockedAsDefault": false
      }
    ]
    ```

=== "Text"

    ```text
    TagId                                 TagName
    ------------------------------------  --------------
    def61080-111c-4aea-b72f-5b60e516e36c  Some label
    ```

=== "CSV"

    ```csv
    AcceptMessagesOnlyFromSendersOrMembers,AutoDelete,BlockDelete,BlockEdit,ComplianceFlags,ContainsSiteLabel,DisplayName,HasRetentionAction,IsEventTag,RequireSenderAuthenticationEnabled,SuperLock,TagDuration,TagId,TagName,TagRetentionBasedOn,UnlockedAsDefault
    ,1,1,,1,,,1,,,,2555,def61080-111c-4aea-b72f-5b60e516e36c,Some label,CreationAgeInDays,
    ```
    
=== "Markdown"

    ```md
    # spo web retentionlabel list --webUrl "https://reshmeeauckloo.sharepoint.com/sites/Company311"

    Date: 4/11/2023

    Property | Value
    ---------|-------
    AcceptMessagesOnlyFromSendersOrMembers | false
    AutoDelete | true
    BlockDelete | true
    BlockEdit | false
    ComplianceFlags | 1
    ContainsSiteLabel | false
    DisplayName | 
    HasRetentionAction | true
    IsEventTag | false
    RequireSenderAuthenticationEnabled | false
    SuperLock | false
    TagDuration | 2555
    TagId | def61080-111c-4aea-b72f-5b60e516e36c
    TagName | Some Label
    TagRetentionBasedOn | CreationAgeInDays
    UnlockedAsDefault | false
    ```
