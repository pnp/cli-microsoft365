# spo serviceprincipal grant add

Grants the service principal permission to the specified API

## Usage

```sh
m365 spo serviceprincipal grant add [options]
```

## Alias

```sh
m365 spo sp grant add
```

## Options

`-r, --resource <resource>`
: The name of the resource for which permissions should be granted.

`-s, --scope <scope>`
: The name of the permission that should be granted.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator.

## Examples

Grant the service principal permission to read email using the Microsoft Graph

```sh
m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Mail.Read'
```

Grant the service principal permission to a custom API

```sh
m365 spo serviceprincipal grant add --resource 'contoso-api' --scope 'user_impersonation'
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
      "Scope": "Mail.Read"
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
    Scope           : Mail.Read
    ```

=== "CSV"

    ```csv
    ClientId,ConsentType,IsDomainIsolated,ObjectId,PackageName,Resource,ResourceId,Scope
    6004a642-185c-479a-992a-15d1c23e2229,AllPrincipals,,QqYEYFwYmkeZKhXRwj4iKRcAa6TiIbFNvGnKY1dqONY,,Microsoft Graph,a46b0017-21e2-4db1-bc69-ca63576a38d6,Mail.Read
    ```

=== "Markdown"

    ```md
    # spo serviceprincipal grant add --resource "Microsoft Graph" --scope "Mail.Read"

    Date: 5/7/2023

    ## 4WtBzD8u5kW-sYuikIWL_8ZYTP5mJB1LnC6OT4Ibr94

    Property | Value
    ---------|-------
    ClientId | cc416be1-2e3f-45e6-beb1-8ba290858bff
    ConsentType | AllPrincipals
    IsDomainIsolated | false
    ObjectId | 4WtBzD8u5kW-sYuikIWL\_8ZYTP5mJB1LnC6OT4Ibr94
    Resource | Microsoft Graph
    ResourceId | fe4c58c6-2466-4b1d-9c2e-8e4f821bafde
    Scope | Mail.Read
    ```
