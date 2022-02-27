# teams app install

Installs a Microsoft Teams team app from the catalog in the specified team or for the specified user

## Usage

```sh
m365 teams app install [options]
```

## Options

`--appId <appId>`
: The ID of the app to install

`--teamId [teamId]`
: The ID of the Microsoft Teams team to which to install the app

`--userId [userId]`
: The ID of the user for who to install the app. Specify either `userId` or `userName` to install a personal app for a user.

`--userName [userName]`
: The UPN of the user for who to install the app. Specify either `userId` or `userName` to install a personal app for a user.

--8<-- "docs/cmd/_global.md"

## Remarks

The `appId` has to be the ID of the app from the Microsoft Teams App Catalog. Do not use the ID from the manifest of the zip app package. Use the [teams app list](./app-list.md) command to get this ID instead.

## Examples

Install an app from the catalog in a Microsoft Teams team

```sh
m365 teams app install --appId 4440558e-8c73-4597-abc7-3644a64c4bce --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```

Install a personal app for the user specified using their user name

```sh
m365 teams app install --appId 4440558e-8c73-4597-abc7-3644a64c4bce --userName steve@contoso.com
```

Install a personal app for the user specified using their ID

```sh
m365 teams app install --appId 4440558e-8c73-4597-abc7-3644a64c4bce --userId 2609af39-7775-4f94-a3dc-0dd67657e900
```
