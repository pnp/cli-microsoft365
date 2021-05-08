# aad app role list

Gets Azure AD app registration roles

## Usage

```sh
m365 aad app role list [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration for which to retrieve roles. Specify either `appId`, `appObjectId` or `appName`

`--appObjectId [appObjectId]`
: Object ID of the Azure AD application registration for which to retrieve roles. Specify either `appId`, `appObjectId` or `appName`

`--appName [appName]`
: Name of the Azure AD application registration for which to retrieve roles. Specify either `appId`, `appObjectId` or `appName`

--8<-- "docs/cmd/_global.md"

## Remarks

For best performance use the `appObjectId` option to reference the Azure AD application registration for which to retrieve roles. If you use `appId` or `appName`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

## Examples

Get roles for the Azure AD application registration specified by its object ID

```sh
m365 aad app role list --appObjectId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Get roles for the Azure AD application registration specified by its app (client) ID

```sh
m365 aad app role list --appId e75be2e1-0204-4f95-857d-51a37cf40be8
```

Get roles for the Azure AD application registration specified by its name

```sh
m365 aad app role list --appName "My app"
```
