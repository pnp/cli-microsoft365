# spo web list

Lists subsites of the specified site

## Usage

```sh
m365 spo web list [options]
```

## Options

`-u, --url <url>`
: URL of the parent site for which to retrieve the list of subsites

--8<-- "docs/cmd/_global.md"

## Examples

Return all subsites from site _https://contoso.sharepoint.com/_

```sh
m365 spo web list --url https://contoso.sharepoint.com
```

## Response

=== "JSON"

    ```json
    [
      {
        "AllowRssFeeds": true,
        "AlternateCssUrl": "",
        "AppInstanceId": "00000000-0000-0000-0000-000000000000",
        "ClassicWelcomePage": null,
        "Configuration": 0,
        "Created": "2022-11-05T14:07:51",
        "CurrentChangeToken": {
          "StringValue": "1;2;b60137df-c3dc-4984-9def-8edcf7c98ab9;638032561833570000;715587227"
        },
        "CustomMasterUrl": "/subsite/_catalogs/masterpage/seattle.master",
        "Description": "Subsite",
        "DesignPackageId": "00000000-0000-0000-0000-000000000000",
        "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
        "EnableMinimalDownload": true,
        "FooterEmphasis": 0,
        "FooterEnabled": false,
        "FooterLayout": 0,
        "HeaderEmphasis": 0,
        "HeaderLayout": 0,
        "HideTitleInHeader": false,
        "HorizontalQuickLaunch": false,
        "Id": "b60137df-c3dc-4984-9def-8edcf7c98ab9",
        "IsEduClass": false,
        "IsEduClassProvisionChecked": false,
        "IsEduClassProvisionPending": false,
        "IsHomepageModernized": false,
        "IsMultilingual": false,
        "IsRevertHomepageLinkHidden": false,
        "Language": 1033,
        "LastItemModifiedDate": "2022-11-05T14:08:03Z",
        "LastItemUserModifiedDate": "2022-11-05T14:08:03Z",
        "LogoAlignment": 0,
        "MasterUrl": "/subsite/_catalogs/masterpage/seattle.master",
        "MegaMenuEnabled": false,
        "NavAudienceTargetingEnabled": false,
        "NoCrawl": false,
        "ObjectCacheEnabled": false,
        "OverwriteTranslationsOnChange": false,
        "ResourcePath": {
          "DecodedUrl": "https://contoso.sharepoint.com/subsite"
        },
        "QuickLaunchEnabled": true,
        "RecycleBinEnabled": true,
        "SearchScope": 0,
        "ServerRelativeUrl": "/subsite",
        "SiteLogoUrl": null,
        "SyndicationEnabled": true,
        "TenantAdminMembersCanShare": 0,
        "Title": "Subsite",
        "TreeViewEnabled": false,
        "UIVersion": 15,
        "UIVersionConfigurationEnabled": false,
        "Url": "https://contoso.sharepoint.com/subsite",
        "WebTemplate": "STS",
        "WelcomePage": "SitePages/Home.aspx"
      }
    ]
    ```

=== "Text"

    ```text
    Title    Url                                      Id
    -------  ---------------------------------------  ------------------------------------
    Subsite  https://contoso.sharepoint.com/subsite   b60137df-c3dc-4984-9def-8edcf7c98ab9
    ```

=== "CSV"

    ```csv
    Title,Url,Id
    Subsite,https://contoso.sharepoint.com/subsite,b60137df-c3dc-4984-9def-8edcf7c98ab9
    ```
