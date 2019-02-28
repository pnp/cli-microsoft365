# graph teams app publish

Publishes Teams app to the organization's app catalog

## Usage

```sh
graph teams app publish [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --filePath <filePath>`|Absolute or relative path to the Teams manifest zip file to add to the app catalog
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To publish Microsoft Teams apps, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only publish a Teams app as a global administrator.

## Examples

Add the _teams-manifest.zip_ file to the organization's app catalog

```sh
graph teams app publish --filePath ./teams-manifest.zip
```