# graph schemaextension get

Gets the properties of the specified schema extension definition

## Usage

```sh
m365 graph schemaextension get [options]
```

## Options

`-i, --id <id>`
: The unique identifier for the schema extension definition

--8<-- "docs/cmd/_global.md"

## Remarks

To get properties of a schema extension definition, you have to pass the ID of the schema
extension.

## Examples

Gets properties of a schema extension definition with ID domain_myExtension

```sh
m365 graph schemaextension get --id domain_myExtension 
```
