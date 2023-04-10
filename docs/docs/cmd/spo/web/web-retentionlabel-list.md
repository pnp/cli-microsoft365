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
    TagId,TagName
    def61080-111c-4aea-b72f-5b60e516e36c,Some label,true
    ```
    
=== "Markdown"

    ```md
    # m365 spo web retentionlabel list --webUrl 'https://contoso.sharepoint.com/sites/sales'
    
    Date: 4/10/2023    

    ## Some label (def61080-111c-4aea-b72f-5b60e516e36cm3)

    Property | Value
    TagId | def61080-111c-4aea-b72f-5b60e516e36c
    TagName | Some label
    ```
