# teams app uninstall

Uninstalls an app from a Microsoft Team team

## Usage

```sh
m365 teams app uninstall [options]
```

## Options

`--appId <appId>`
: The unique id of the app instance installed in the Microsoft Teams team

`--teamId <teamId>`
: The ID of the Microsoft Teams team from which to uninstall the app

`--confirm`
: Don't prompt for confirmation when uninstalling the app

--8<-- "docs/cmd/_global.md"

## Remarks

The `appId` has to be the id the app instance installed in the Microsoft Teams team.
Do not use the ID from the manifest of the zip app package or the id from the Microsoft Teams App Catalog.

## Examples

Uninstall an app from a Microsoft Teams team

```sh
m365 teams app uninstall --appId YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY= --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```