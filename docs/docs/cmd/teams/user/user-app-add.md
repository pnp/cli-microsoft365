# teams user app add

Install an app in the personal scope of the specified user

## Usage

```sh
m365 teams user app add [options]
```

## Options

`--appId <appId>`
: The ID of the app to install

`--userId <userId>`
: The ID of the user to install the app for

--8<-- "docs/cmd/_global.md"

## Remarks

The `appId` has to be the ID of the app from the Microsoft Teams App Catalog. Do not use the ID from the manifest of the zip app package. Use the [teams app list](../app/app-list.md) command to get this ID.

## Examples

Install an app from the catalog for the specified user

```sh
m365 teams user app add --appId 4440558e-8c73-4597-abc7-3644a64c4bce --userId 2609af39-7775-4f94-a3dc-0dd67657e900
```
