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

`-i, --id [id]`
: ID of the permission request to approve

`--all`
: set, to approva all pending permission requests

`--resource [resource]`
: The resource of the permissions requests to approve

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The admin role that's required to approve permissions depends on the API. To approve permissions to any of the third-party APIs registered in the tenant, the application administrator role is sufficient. To approve permissions for Microsoft Graph or any other Microsoft API, the Global Administrator role is required.

The permission request you want to approve is denoted using its `ID`. You can retrieve it using the [spo serviceprincipal permissionrequest list](./serviceprincipal-permissionrequest-list.md) command.

## Examples

Approve permission request with id

```sh
m365 spo serviceprincipal permissionrequest approve --id 4dc4c043-25ee-40f2-81d3-b3bf63da7538
```

Approve all permission request

```sh
m365 spo serviceprincipal permissionrequest approve --all
```

Approve all permission request from a specific resource

```sh
m365 spo serviceprincipal permissionrequest approve --resource "Microsoft Graph"
```

## Response

=== "JSON"

    ```json
    {
      "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d",
      "ConsentType": "AllPrincipals",
      "IsDomainIsolated": false,
      "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM",
      "PackageName": null,
      "Resource": "Microsoft Graph",
      "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903",
      "Scope": "User.Read.All"
    }
    ```

=== "Text"

    ```text
    ClientId        : 90a2c08e-e786-4100-9ea9-36c261be6c0d
    ConsentType     : AllPrincipals
    IsDomainIsolated: false
    ObjectId        : jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM
    PackageName     : null
    Resource        : Microsoft Graph
    ResourceId      : d6afc295-0a08-4777-a095-10bcb7e23903
    Scope           : User.Read.All
    ```

=== "CSV"

    ```csv
    ClientId,ConsentType,IsDomainIsolated,ObjectId,PackageName,Resource,ResourceId,Scope
    90a2c08e-e786-4100-9ea9-36c261be6c0d,AllPrincipals,,jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM,,Microsoft Graph,d6afc295-0a08-4777-a095-10bcb7e23903,User.Read.All
    ```

### `all`, `resource` response

When we make use of the option `all` or `resource` the response will differ.

=== "JSON"

    ```json
    [
      {
        "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d",
        "ConsentType": "AllPrincipals",
        "IsDomainIsolated": false,
        "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM",
        "PackageName": null,
        "Resource": "Microsoft Graph",
        "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903",
        "Scope": "User.Read.All"
      },
      {
        "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d",
        "ConsentType": "AllPrincipals",
        "IsDomainIsolated": false,
        "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM",
        "PackageName": null,
        "Resource": "Microsoft Graph",
        "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903",
        "Scope": "Sites.Read.All"
      }
    ]
    ```

=== "Text"

    ```text
    ClientId                              ConsentType    IsDomainIsolated  ObjectId                                     PackageName  Resource         ResourceId                            Scope
    ------------------------------------  -------------  ----------------  -------------------------------------------  -----------  ---------------  ------------------------------------  -----------------------
    90a2c08e-e786-4100-9ea9-36c261be6c0d  AllPrincipals  false             jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM  null         Microsoft Graph  d6afc295-0a08-4777-a095-10bcb7e23903  User.Read.All
    90a2c08e-e786-4100-9ea9-36c261be6c0d  AllPrincipals  false             jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM  null         Microsoft Graph  d6afc295-0a08-4777-a095-10bcb7e23903  Sites.Read.All
    ```

=== "CSV"

    ```csv
    ClientId,ConsentType,IsDomainIsolated,ObjectId,PackageName,Resource,ResourceId,Scope
    90a2c08e-e786-4100-9ea9-36c261be6c0d,AllPrincipals,,jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM,,Microsoft Graph,d6afc295-0a08-4777-a095-10bcb7e23903,User.Read.All
    90a2c08e-e786-4100-9ea9-36c261be6c0d,AllPrincipals,,jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM,,Microsoft Graph,d6afc295-0a08-4777-a095-10bcb7e23903,Sites.Read.All
    ```
