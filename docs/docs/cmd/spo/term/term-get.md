# spo term get

Gets information about the specified taxonomy term

## Usage

```sh
m365 spo term get [options]
```

## Options

`-u, --webUrl [webUrl]`
: If specified, allows you to get a term from the tenant term store as well as the sitecollection specific term store. Defaults to the tenant admin site.

`-i, --id [id]`
: ID of the term to retrieve. Specify `name` or `id` but not both.

`-n, --name [name]`
: Name of the term to retrieve. Specify `name` or `id` but not both.

`--termGroupId [termGroupId]`
: ID of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both.

`--termGroupName [termGroupName]`
: Name of the term group to which the term set belongs. Specify `termGroupId` or `termGroupName` but not both.

`--termSetId [termSetId]`
: ID of the term set to which the term belongs. Specify `termSetId` or `termSetName` but not both.

`--termSetName [termSetName]`
: Name of the term set to which the term belongs. Specify `termSetId` or `termSetName` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

When retrieving term by its ID, it's sufficient to specify just the ID. When retrieving it by its name however, you need to specify the parent term group and term set using either their names or IDs.

!!! important
    To use this command without the --webUrl option you have to have permissions to access the tenant admin site.
    
When using the `--webUrl` option you can connect to the term store with limited permissions, and do not need the SharePoint Adminstrator role. You need be a site visitor or more. It allows you to get a term from the tenant term store as well as a term from the sitecollection term store.

## Examples

Get information about a taxonomy term using its ID from the specified sitecollection.

```sh
m365 spo term get --webUrl https://contoso.sharepoint.com/sites/project-x --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term using its ID.

```sh
m365 spo term get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term using its name, retrieving the parent term group and term set using their names.

```sh
m365 spo term get --name IT --termGroupName People --termSetName Department
```

Get information about a taxonomy term using its name, retrieving the parent term group and term set using their IDs.

```sh
m365 spo term get --name IT --termGroupId 5c928151-c140-4d48-aab9-54da901c7fef --termSetId 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f
```

## Response

=== "JSON"

    ```json
    {
      "CreatedDate": "2021-07-07T09:42:02.283Z",
      "Id": "2b5c71a6-d72b-49a8-a3bf-d80636d85b44",
      "LastModifiedDate": "2021-07-07T09:42:02.283Z",
      "Name": "IT",
      "CustomProperties": {},
      "CustomSortOrder": null,
      "IsAvailableForTagging": true,
      "Owner": "NT Service\\SPTimerV4",
      "Description": "",
      "IsDeprecated": false,
      "IsKeyword": false,
      "IsPinned": false,
      "IsPinnedRoot": false,
      "IsReused": false,
      "IsRoot": true,
      "IsSourceTerm": true,
      "LocalCustomProperties": {},
      "MergedTermIds": [],
      "PathOfTerm": "IT",
      "TermsCount": 1
    }
    ```

=== "Text"

    ```text
    CreatedDate          : 2021-07-07T09:42:02.283Z
    CustomProperties     : {}
    CustomSortOrder      : null
    Description          :
    Id                   : 2b5c71a6-d72b-49a8-a3bf-d80636d85b44
    IsAvailableForTagging: true
    IsDeprecated         : false
    IsKeyword            : false
    IsPinned             : false
    IsPinnedRoot         : false
    IsReused             : false
    IsRoot               : true
    IsSourceTerm         : true
    LastModifiedDate     : 2021-07-07T09:42:02.283Z
    LocalCustomProperties: {}
    MergedTermIds        : []
    Name                 : IT
    Owner                : NT Service\SPTimerV4
    PathOfTerm           : IT
    TermsCount           : 1
    ```

=== "CSV"

    ```csv
    CreatedDate,Id,LastModifiedDate,Name,IsAvailableForTagging,Owner,Description,IsDeprecated,IsKeyword,IsPinned,IsPinnedRoot,IsReused,IsRoot,IsSourceTerm,PathOfTerm,TermsCount
    2021-07-07T09:42:02.283Z,2b5c71a6-d72b-49a8-a3bf-d80636d85b44,2021-07-07T09:42:02.283Z,IT,1,NT Service\SPTimerV4,,,,,,,1,1,IT,1
    ```

=== "Markdown"

    ```md
    # spo term get --termGroupName "People" --termSetName "Department" --name "IT"

    Date: 5/8/2023

    ## IT (2b5c71a6-d72b-49a8-a3bf-d80636d85b44)

    Property | Value
    ---------|-------
    CreatedDate | 2021-07-07T09:42:02.283Z
    Id | 2b5c71a6-d72b-49a8-a3bf-d80636d85b44
    LastModifiedDate | 2021-07-07T09:42:02.283Z
    Name | IT
    IsAvailableForTagging | true
    Owner | NT Service\SPTimerV4
    Description |
    IsDeprecated | false
    IsKeyword | false
    IsPinned | false
    IsPinnedRoot | false
    IsReused | false
    IsRoot | true
    IsSourceTerm | true
    PathOfTerm | IT
    TermsCount | 1
    ```
