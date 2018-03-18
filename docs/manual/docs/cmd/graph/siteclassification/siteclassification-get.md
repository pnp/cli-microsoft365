# graph siteclassification get

Gets information about the Office 365 Tenant SiteClassification

## Usage

```sh
graph siteclassification get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To get information about a Office 365 Tenant SiteClassification, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

Gets information about the Office 365 Tenant SiteClassification

```sh
graph siteclassification get
```