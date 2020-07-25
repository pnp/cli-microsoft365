# graph schemaextension add

Creates a Microsoft Graph schema extension

## Usage

```sh
m365 graph schemaextension add [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The unique identifier for the schema extension definition

`-d, --description [description]`
: Description of the schema extension

`--owner <owner>`
: The Id ID the Azure AD application that is the owner of the schema extension

`-t, --targetTypes <targetTypes>`
: Comma-separated list of Microsoft Graph resource types the schema extension targets

`-p, --properties`
: The collection of property names and types that make up the schema extension definition formatted as a JSON string

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To create a schema extension, you have to specify a unique ID for the schema extension. You can assign a value in one of two ways:

- concatenate the name of one of your verified domains with a name for the schema extension to form a unique string in this format, `{domainName}_{schemaName}`, eg. `contoso_mySchema`.

    NOTE: Only verified domains under the following top-level domains are supported: .com, .net, .gov, .edu or .org.

- Provide a schema name, and let Microsoft Graph use that schema name to complete the id assignment in this format: `ext{8-random-alphanumeric-chars}_{schema-name}`, eg. `extkvbmkofy_mySchema`.

The schema extension ID cannot be changed after creation.

The schema extension owner is the ID of the Azure AD application that is the owner of the schema extension. Once set, this property is read-only and cannot be changed.

The target types are the set of Microsoft Graph resource types (that support schema extensions) that this schema extension definition can be applied to. This option is specified as a comma-separated list

When specifying the JSON string of properties on Windows, you have to escape double quotes in a specific way. Considering the following value for the _properties_ option: `{"Foo":"Bar"}`,
you should specify the value as <code>\`"{""Foo"":""Bar""}"\`</code>.
In addition, when using PowerShell, you should use the `--%` argument.

## Examples

Create a schema extension

```sh
m365 graph schemaextension add --id MySchemaExtension --description "My Schema Extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
```

Create a schema extension with a verified domain

```sh
m365 graph schemaextension add --id contoso_MySchemaExtension --description "My Schema Extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
```

Create a schema extension in PowerShell

```PowerShell
graph schemaextension add --id contoso_MySchemaExtension --description "My Schema Extension" --targetTypes Group --owner "62375ab9-6b52-47ed-826b-58e47e0e304b" --properties --% \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
```