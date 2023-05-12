# spo term group list

Lists taxonomy term groups

## Usage

```sh
m365 spo term group list [options]
```

## Options

`-u, --webUrl [webUrl]`
: If specified, allows you to list term groups from the tenant term store as well as the sitecollection specific term store. Defaults to the tenant admin site.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command without the --webUrl option you have to have permissions to access the tenant admin site.

When using the `--webUrl` option you can connect to the term store with limited permissions, and do not need the SharePoint Adminstrator role. You need to be a site visitor or more. It allows you to list term groups from the tenant term store as well as term groups from the sitecollection term store.

## Examples

List taxonomy term groups.

```sh
m365 spo term group list
```

List taxonomy term groups from the specified sitecollection.

```sh
m365 spo term group list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    [
      {
        "_ObjectType_": "SP.Taxonomy.TermGroup",
        "_ObjectIdentity_": "5522b1a0-b01a-2000-4160-d04eee2e977f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g8=",
        "CreatedDate": "2019-09-03T06:41:32.070Z",
        "Id": "4aa9cd7a-d275-4fda-84b2-b83085dcee0f",
        "LastModifiedDate": "2019-09-03T06:41:32.070Z",
        "Name": "People",
        "Description": "",
        "IsSiteCollectionGroup": false,
        "IsSystemGroup": false
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Name
    ------------------------------------  -----------------------------------------------------------
    4aa9cd7a-d275-4fda-84b2-b83085dcee0f  People
    ```

=== "CSV"

    ```csv
      _ObjectType_,_ObjectIdentity_,CreatedDate,Id,LastModifiedDate,Name,Description,IsSiteCollectionGroup,IsSystemGroup
      SP.Taxonomy.TermGroup,7522b1a0-804d-2000-41be-bba084eae72f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g8=,2019-09-03T06:41:32.070Z,4aa9cd7a-d275-4fda-84b2-b83085dcee0f,2019-09-03T06:41:32.070Z,People,,,
    ```

=== "Markdown"

    ```md
    # spo term group list

    Date: 5/9/2023

    ## People (4aa9cd7a-d275-4fda-84b2-b83085dcee0f)

    Property | Value
    ---------|-------
    \_ObjectType\_ | SP.Taxonomy.TermGroup
    \_ObjectIdentity\_ | 8c22b1a0-200f-2000-3976-694ba2bf8a01\|fec14c62-7c3b-481b-851b-c80d7802b224:gr:aCf0Cz4D9UOS7+b/OlUY5XrNqUp10tpPhLK4MIXc7g8=
    CreatedDate | 2019-09-03T06:41:32.070Z
    Id | 4aa9cd7a-d275-4fda-84b2-b83085dcee0f
    LastModifiedDate | 2019-09-03T06:41:32.070Z
    Name | People
    Description |
    IsSiteCollectionGroup | false
    IsSystemGroup | false
    ```
