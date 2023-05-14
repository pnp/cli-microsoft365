# spo term group get

Gets information about the specified taxonomy term group

## Usage

```sh
m365 spo term group get [options]
```

## Options

`-u, --webUrl [webUrl]`
: If specified, allows you to get a term group from the tenant term store as well as the sitecollection specific term store. Defaults to the tenant admin site.

`-i, --id [id]`
: ID of the term group to retrieve. Specify `name` or `id` but not both.

`-n, --name [name]`
: Name of the term group to retrieve. Specify `name` or `id` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command without the `--webUrl` option you have to have permissions to access the tenant admin site.

When using the `--webUrl` option you can connect to the term store with limited permissions, and do not need the SharePoint Adminstrator role. You need be a site visitor or more. It allows you to get a term group from the tenant term store as well as a term group from the sitecollection term store.

## Examples

Get information about a taxonomy term group using its ID.

```sh
m365 spo term group get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb
```

Get information about a taxonomy term group using its name.

```sh
m365 spo term group get --name PnPTermSets
```

Get information about a taxonomy term group using its ID from the specified sitecollection.

```sh
m365 spo term group get --id 0e8f395e-ff58-4d45-9ff7-e331ab728beb --webUrl https://contoso.sharepoint.com/sites/project-x
```

## Response

=== "JSON"

    ```json
    {
      "CreatedDate": "2019-09-03T06:41:32.070Z",
      "Id": "0e8f395e-ff58-4d45-9ff7-e331ab728beb",
      "LastModifiedDate": "2019-09-03T06:41:32.070Z",
      "Name": "PnPTermSets",
      "Description": "",
      "IsSiteCollectionGroup": false,
      "IsSystemGroup": false
    }
    ```

=== "Text"

    ```text
    CreatedDate          : 2019-09-03T06:41:32.070Z
    Description          :
    Id                   : 0e8f395e-ff58-4d45-9ff7-e331ab728beb
    IsSiteCollectionGroup: false
    IsSystemGroup        : false
    LastModifiedDate     : 2019-09-03T06:41:32.070Z
    Name                 : PnPTermSets
    ```

=== "CSV"

    ```csv
    CreatedDate,Id,LastModifiedDate,Name,Description,IsSiteCollectionGroup,IsSystemGroup
    2019-09-03T06:41:32.070Z,0e8f395e-ff58-4d45-9ff7-e331ab728beb,2019-09-03T06:41:32.070Z,PnPTermSets,,,
    ```

=== "Markdown"

    ```md
    # spo term group get --id "0e8f395e-ff58-4d45-9ff7-e331ab728beb"

    Date: 5/14/2023

    ## PnPTermSets (0e8f395e-ff58-4d45-9ff7-e331ab728beb)

    Property | Value
    ---------|-------
    CreatedDate | 2019-09-03T06:41:32.070Z
    Id | 0e8f395e-ff58-4d45-9ff7-e331ab728beb
    LastModifiedDate | 2019-09-03T06:41:32.070Z
    Name | PnPTermSets
    Description |
    IsSiteCollectionGroup | false
    IsSystemGroup | false
    ```
