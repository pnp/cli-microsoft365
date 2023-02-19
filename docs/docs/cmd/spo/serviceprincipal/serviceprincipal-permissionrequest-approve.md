# spo serviceprincipal permissionrequest approve

Approves the specified permission request

## Usage

```sh
m365 spo serviceprincipal permissionrequest approve [options]
```

## Alias

```sh
m365 spo sp permissionrequest approve
```

## Options

`-i, --id <id>`
: ID of the permission request to approve.

`--all`
: approve all pending permission requests.

`--resource [resource]`
: The resource of the permissions requests to approve.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The admin role that's required to approve permissions depends on the API. To approve permissions to any of the third-party APIs registered in the tenant, the application administrator role is sufficient. To approve permissions for Microsoft Graph or any other Microsoft API, the Global Administrator role is required.

The permission request you want to approve is denoted using its `ID`. You can retrieve it using the [spo serviceprincipal permissionrequest list](./serviceprincipal-permissionrequest-list.md) command.

## Examples

Approve permission request

```sh
m365 spo serviceprincipal permissionrequest approve --id 4dc4c043-25ee-40f2-81d3-b3bf63da7538
```

## Response

=== "JSON"

    ```json
    {
      "ClientId": "6004a642-185c-479a-992a-15d1c23e2229",
      "ConsentType": "AllPrincipals",
      "IsDomainIsolated": false,
      "ObjectId": "QqYEYFwYmkeZKhXRwj4iKRcAa6TiIbFNvGnKY1dqONY",
      "PackageName": null,
      "Resource": "Microsoft Graph",
      "ResourceId": "a46b0017-21e2-4db1-bc69-ca63576a38d6",
      "Scope": "Reports.Read.All"
    }
    ```

=== "Text"

    ```text
    ClientId        : 6004a642-185c-479a-992a-15d1c23e2229
    ConsentType     : AllPrincipals
    IsDomainIsolated: false
    ObjectId        : QqYEYFwYmkeZKhXRwj4iKRcAa6TiIbFNvGnKY1dqONY
    PackageName     : null
    Resource        : Microsoft Graph
    ResourceId      : a46b0017-21e2-4db1-bc69-ca63576a38d6
    Scope           : Directory.ReadWrite.All
    ```

=== "CSV"

    ```csv
    ClientId,ConsentType,IsDomainIsolated,ObjectId,PackageName,Resource,ResourceId,Scope
    6004a642-185c-479a-992a-15d1c23e2229,AllPrincipals,false,QqYEYFwYmkeZKhXRwj4iKRcAa6TiIbFNvGnKY1dqONY,null,Microsoft Graph,a46b0017-21e2-4db1-bc69-ca63576a38d6,Directory.ReadWrite.All
    ```
