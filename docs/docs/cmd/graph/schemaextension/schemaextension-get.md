# graph schemaextension get

Gets the properties of the specified schema extension definition

## Usage

```sh
m365 graph schemaextension get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The unique identifier for the schema extension definition

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To get properties of a schema extension definition, you have to pass the ID of the schema
extension.

## Examples

Gets properties of a schema extension definition with ID domain_myExtension

```sh
m365 graph schemaextension get --id domain_myExtension 
```
