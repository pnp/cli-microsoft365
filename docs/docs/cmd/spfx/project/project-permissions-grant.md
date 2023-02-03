# spfx project permissions grant

Grant API permissions defined in the current SPFx project

## Usage

```sh
m365 spfx project permissions grant [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    Run this command in the folder where the project is located from where you want to grant the permissions.

This command grant the permissions defined in: _package-solution.json_.

## Examples

Grant API permissions requested in the current SPFx project

```sh
m365 spfx project permissions grant
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
