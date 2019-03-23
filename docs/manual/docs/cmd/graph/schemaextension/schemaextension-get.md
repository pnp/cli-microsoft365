# graph schemaextension get

Gets information about a Microsoft Graph schema extension

## Usage

```sh
graph schemaextension get <options>
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The unique identifier for the schema extension definition
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging


!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To get information about a schema extension, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

To get information about a schema extension, you have to specify the unique ID of the schema extension

## Examples

Gets information about a schema extension with ID MySchemaExtension
```sh
  graph schemaextension get --id MySchemaExtension
```