# graph teams app remove

Removes a Teams app from the oranization's app catalog

## Usage

```sh
graph teams app remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to upgrade
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

### Remarks

To remove a Teams app from your organzation's app catalog, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only remove a Teams app as a global administrator.

## Examples

Remove a Teams app

```sh
graph teams app remove --id 83cece1e-938d-44a1-8b86-918cf6151957
```