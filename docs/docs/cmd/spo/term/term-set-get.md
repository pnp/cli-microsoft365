# spo term set get

Gets information about the specified taxonomy term set

## Usage

```sh
m365 spo term set get [options]
```

## Options

`-u, --webUrl [webUrl]`
: If specified, allows you to get a term set from the tenant term store as well as the sitecollection specific term store. Defaults to the tenant admin site.

`-i, --id [id]`
: ID of the term set to retrieve. Specify `name` or `id` but not both.

`-n, --name [name]`
: Name of the term set to retrieve. Specify `name` or `id` but not both.

`--termGroupId [termGroupId]`
: ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both.

`--termGroupName [termGroupName]`
: Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command without the --webUrl option you have to have permissions to access the tenant admin site.

When using the `--webUrl` option you can connect to the term store with limited permissions, and do not need the SharePoint Adminstrator role. You need be a site visitor or more. It allows you to get a term set from the tenant term store as well as a term set from the sitecollection term store.

## Examples

Get information about a taxonomy term set using its ID.

```sh
m365 spo term set get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termGroupName PnPTermSets
```

Get information about a taxonomy term set using its name.

```sh
m365 spo term set get --name PnP-Organizations --termGroupId 0a099ee9-e231-4ae9-a5b6-d7f94a0d241d
```

Get information about a taxonomy term set using its ID from the specified sitecollection.

```sh
m365 spo term set get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --termGroupName PnPTermSets --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    {
      "CreatedDate": "2019-09-03T06:41:32.110Z",
      "Id": "0e8f395e-ff58-4d45-9ff7-e331ab728beb",
      "LastModifiedDate": "2019-09-03T06:41:32.260Z",
      "Name": "PnP-Organizations",
      "CustomProperties": {
        "SearchCenterNavVer": "15"
      },
      "CustomSortOrder": null,
      "IsAvailableForTagging": true,
      "Owner": "",
      "Contact": "",
      "Description": "",
      "IsOpenForTermCreation": true,
      "Names": {
        "1033": "PnP-Organizations"
      },
      "Stakeholders": []
    }
    ```

=== "Text"

    ```text
    Contact              :
    CreatedDate          : 2019-09-03T06:41:32.110Z
    CustomProperties     : {"SearchCenterNavVer":"15"}
    CustomSortOrder      : null
    Description          :
    Id                   : 0e8f395e-ff58-4d45-9ff7-e331ab728beb
    IsAvailableForTagging: true
    IsOpenForTermCreation: true
    LastModifiedDate     : 2019-09-03T06:41:32.260Z
    Name                 : PnP-Organizations
    Names                : {"1033":"PnP-Organizations"}
    Owner                :
    Stakeholders         : []
    ```

=== "CSV"

    ```csv
    CreatedDate,Id,LastModifiedDate,Name,IsAvailableForTagging,Owner,Contact,Description,IsOpenForTermCreation
    2019-09-03T06:41:32.110Z,0e8f395e-ff58-4d45-9ff7-e331ab728beb,2019-09-03T06:41:32.260Z,PnP-Organizations,1,,,,1
    ```

=== "Markdown"

    ```md
    # spo term set get --name "PnP-Organizations" --termGroupName "PnPTermSets"

    Date: 5/9/2023

    ## PnP-Organizations (0e8f395e-ff58-4d45-9ff7-e331ab728beb)

    Property | Value
    ---------|-------
    CreatedDate | 2019-09-03T06:41:32.110Z
    Id | 0e8f395e-ff58-4d45-9ff7-e331ab728beb
    LastModifiedDate | 2019-09-03T06:41:32.260Z
    Name | PnP-Organizations
    IsAvailableForTagging | true
    Owner |
    Contact |
    Description |
    IsOpenForTermCreation | true
    ```
