# graph schemaextension remove

Removes specified Microsoft Graph schema extension

## Usage

```sh
m365 graph schemaextension remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The unique identifier for the schema extension definition

`--confirm`
: Don't prompt for confirming removing the specified schema extension

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To remove specified schema extension definition, you have to pass the ID of the schema
extension.

## Examples

Removes specified Microsoft Graph schema extension with ID domain_myExtension. Will prompt for confirmation

```sh
m365 graph schemaextension remove --id domain_myExtension 
```

Removes specified Microsoft Graph schema extension with ID domain_myExtension without prompt for confirmation

```sh
m365 graph schemaextension remove --id domain_myExtension --confirm
```
