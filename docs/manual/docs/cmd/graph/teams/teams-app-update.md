# graph teams app update

Updates Teams app in the organization's app catalog

## Usage

```sh
graph teams app update [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to update
`-p, --filePath <filePath>`|Absolute or relative path to the Teams manifest zip file to update in the app catalog
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update Microsoft Teams apps, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only update a Teams app as a global administrator.

## Examples

Update the Teams app with ID _83cece1e-938d-44a1-8b86-918cf6151957_ from file _teams-manifest.zip_

```sh
graph teams app update --id 83cece1e-938d-44a1-8b86-918cf6151957 --filePath ./teams-manifest.zip
```