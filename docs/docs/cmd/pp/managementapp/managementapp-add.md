# pp managementapp add

Register management application for Power Platform

## Usage

```sh
m365 pp managementapp add [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration to register as management app. Specify either `appId`, `objectId` or `name`

`--objectId [objectId]`
: Object ID of the Azure AD application registration to register as management app. Specify either `appId`, `objectId` or `name`

`--name [name]`
: Name of the Azure AD application registration to register as management app.. Specify either `appId`, `objectId` or `name`

--8<-- "docs/cmd/_global.md"

## Remarks

To execute this command the first time you will need to be signed in using the default Azure PowerShell App Registration `1950a258-227b-4e31-a9cf-717495945fc2`. Run `m365 login --appId 1950a258-227b-4e31-a9cf-717495945fc2` before executing this command. If you are using your the default CLI for Microsoft 365 you can use `--appId 31359c7f-bd7e-475c-86db-fdb8c937548e` to register.

For best performance use the `appId` option to reference the Azure AD application registration to update. If you use `objectId` or `name`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

## Examples

Register CLI for Microsoft 365 as a management application for the Power Platform

```sh
m365 pp managementapp add --appId 31359c7f-bd7e-475c-86db-fdb8c937548e
```

Register your custom application using the ObjectId as a management application for the Power Platform

```sh
m365 pp managementapp add --objectId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Register your My App as a management application for the Power Platform

```sh
m365 pp managementapp add --name "My app"
```
