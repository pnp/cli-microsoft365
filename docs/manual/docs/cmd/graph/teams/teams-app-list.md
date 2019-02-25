# graph teams app list

Lists apps from the Microsoft Teams app catalog

## Usage

```sh
graph teams app list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-a, --all`|Get apps from your organization's app catalog and the Microsoft Teams store
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list apps in the Microsoft Teams app catalog, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List all Microsoft Teams apps from your organization's app catalog only

```sh
graph teams app list
```

List all apps from the Microsoft Teams app catalog and the Microsoft Teams store

```sh
graph teams app list --all
```