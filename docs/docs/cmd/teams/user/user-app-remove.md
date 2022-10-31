# teams user app remove

Uninstall an app from the personal scope of the specified user

## Usage

```sh
m365 teams user app remove [options]
```

## Options

`--id <id>`
: The unique id of the app instance installed for the user

`--userId <userId>`
: The ID of the user to uninstall the app for

`--confirm`
: Confirm removal of app for user

--8<-- "docs/cmd/_global.md"

## Remarks

The `id` has to be the id of the app instance installed for the user.
Do not use the ID from the manifest of the zip app package or the id from the Microsoft Teams App Catalog.

## Examples

Uninstall an app for the specified user

```sh
m365 teams user app remove --id YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY= --userId 2609af39-7775-4f94-a3dc-0dd67657e900
```

## Response

The command won't return a response on success.
