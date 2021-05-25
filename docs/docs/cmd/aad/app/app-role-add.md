# aad app role add

Adds role to the specified Azure AD app registration

## Usage

```sh
m365 aad app role add [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration to which to add the role. Specify either `appId`, `appObjectId` or `appName`

`--appObjectId [appObjectId]`
: Object ID of the Azure AD application registration to which to add the role. Specify either `appId`, `appObjectId` or `appName`

`--appName [appName]`
: Name of the Azure AD application registration to which to add the role. Specify either `appId`, `appObjectId` or `appName`

`-n, --name <name>`
: Name of the role to add

`-d, --description <description>`
: Description of the role to add

`-m, --allowedMembers <allowedMembers>`
: Types of members that can be added to the group. Allowed values: `usersGroups`, `applications`, `both`

`-c, --claim <claim>`
: Claim value

--8<-- "docs/cmd/_global.md"

## Remarks

For best performance use the `appObjectId` option to reference the Azure AD application registration for which to add the role. If you use `appId` or `appName`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

## Examples

Add role to the Azure AD application registration specified by its object ID

```sh
m365 aad app role add --appObjectId d75be2e1-0204-4f95-857d-51a37cf40be8 --name Managers --description "Managers" --allowedMembers usersGroups --claim managers
```

Add role to the Azure AD application registration specified by its app (client) ID

```sh
m365 aad app role add --appId e75be2e1-0204-4f95-857d-51a37cf40be8 --name Managers --description "Managers" --allowedMembers usersGroups --claim managers
```

Add role to the Azure AD application registration specified by its name

```sh
m365 aad app role add --appName "My app" --name Managers --description "Managers" --allowedMembers usersGroups --claim managers
```
