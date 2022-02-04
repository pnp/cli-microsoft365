# aad app get

Gets an Azure AD app registration

## Usage

```sh
m365 aad app get [options]
```

## Options

`--appId [appId]`
: Application (client) ID of the Azure AD application registration to get. Specify either `appId`, `objectId` or `name`

`--objectId [objectId]`
: Object ID of the Azure AD application registration to get. Specify either `appId`, `objectId` or `name`

`--name [name]`
: Name of the Azure AD application registration to get. Specify either `appId`, `objectId` or `name`

`--save`
: Use to store the information about the created app in a local file

--8<-- "docs/cmd/_global.md"

## Remarks

For best performance use the `objectId` option to reference the Azure AD application registration to get. If you use `appId` or `name`, this command will first need to find the corresponding object ID for that application.

If the command finds multiple Azure AD application registrations with the specified app name, it will prompt you to disambiguate which app it should use, listing the discovered object IDs.

If you want to store the information about the Azure AD app registration, use the `--save` option. This is useful when you build solutions connected to Microsoft 365 and want to easily manage app registrations used with your solution. When you use the `--save` option, after you get the app registration, the command will write its ID and name to the `.m365rc.json` file in the current directory. If the file already exists, it will add the information about the App registration to it if it's not already present, allowing you to track multiple apps. If the file doesn't exist, the command will create it.

## Examples

Get the Azure AD application registration by its app (client) ID

```sh
m365 aad app get --appId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Get the Azure AD application registration by its object ID

```sh
m365 aad app get --objectId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Get the Azure AD application registration by its name

```sh
m365 aad app get --name "My app"
```

Get the Azure AD application registration by its name. Store information about the retrieved app registration in the _.m365rc.json_ file in the current directory.

```sh
m365 aad app get --name "My app" --save
```
