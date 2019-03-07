# graph teams app install

Uninstall an app from a Microsoft Team

## Usage

```sh
graph teams app uninstall [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--appId <appId>`| The unique id of the app instance installed in the Team
`--teamId <teamId>`| The id of the Microsoft Team from which the app has to be uninstalled
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To uninstall an app from a Team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

The appId has to be the id the app instance installed in the Team. Do not use the ID from the manifest of the zip app package or the id from the app catalog.

## Examples

Uninstall an app from a Microsoft Team

```sh
graph teams app uninstall --appId YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY= --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```