# spo term set list

Lists taxonomy term sets from the given term group

## Usage

```sh
m365 spo term set list [options]
```

## Options

`-u, --webUrl [webUrl]`
: If specified, allows you to list term sets from the tenant term store as well as the sitecollection specific term store. Defaults to the tenant admin site.

`--termGroupId [termGroupId]`
: ID of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both.

`--termGroupName [termGroupName]`
: Name of the term group from which to retrieve term sets. Specify `termGroupName` or `termGroupId` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command without the --webUrl option you have to have permissions to access the tenant admin site.

When using the `--webUrl` option you can connect to the term store with limited permissions, and do not need the SharePoint Administrator role. You need to be a site visitor or more. It allows you to list term sets from the tenant term store as well as term sets from the sitecollection term store.

## Examples

List taxonomy term sets from the term group with the given name.

```sh
m365 spo term set list --termGroupName PnPTermSets
```

List taxonomy term sets from the term group with the given ID.

```sh
m365 spo term set list --termGroupId 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

List taxonomy term sets from the specified sitecollection, from the term group with the given name.

```sh
m365 spo term set list --termGroupName PnPTermSets --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    [
      {
        "_ObjectType_": "SP.Taxonomy.TermSet",
        "_ObjectIdentity_": "676eb2a0-80e7-2000-3976-610eb3f5a91e|fec14c62-7c3b-481b-851b-c80d7802b224:se:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g/qydiOUnAdTKTXucEL/+pv",
        "CreatedDate": "2019-09-03T06:41:32.110Z",
        "Id": "8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f",
        "LastModifiedDate": "2019-09-03T06:41:32.260Z",
        "Name": "Department",
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
          "1033": "Department"
        },
        "Stakeholders": []
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Name
    ------------------------------------  ----------
    8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f  Department
    ```

=== "CSV"

    ```csv
    _ObjectType_,_ObjectIdentity_,CreatedDate,Id,LastModifiedDate,Name,IsAvailableForTagging,Owner,Contact,Description,IsOpenForTermCreation
    SP.Taxonomy.TermSet,7c6eb2a0-20fe-2000-47f1-69774ac909f5|fec14c62-7c3b-481b-851b-c80d7802b224:se:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g/qydiOUnAdTKTXucEL/+pv,2019-09-03T06:41:32.110Z,8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f,2019-09-03T06:41:32.260Z,Department,1,,,,1
    ```

=== "Markdown"

    ```md
    # spo term set list --termGroupName "PnPTermSets"

    Date: 5/13/2023

    ## Department (8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f)

    Property | Value
    ---------|-------
    \_ObjectType\_ | SP.Taxonomy.TermSet
    \_ObjectIdentity\_ | 826eb2a0-405f-2000-41be-bfe56d4cdc73\|fec14c62-7c3b-481b-851b-c80d7802b224:se:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g/qydiOUnAdTKTXucEL/+pv
    CreatedDate | 2019-09-03T06:41:32.110Z
    Id | 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f
    LastModifiedDate | 2019-09-03T06:41:32.260Z
    Name | Department
    IsAvailableForTagging | true
    Owner |
    Contact |
    Description |
    IsOpenForTermCreation | true
    ```
