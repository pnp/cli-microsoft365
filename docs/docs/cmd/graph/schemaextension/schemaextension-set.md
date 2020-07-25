# graph schemaextension set

Updates a Microsoft Graph schema extension

## Usage

```sh
m365 graph schemaextension set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The unique identifier for the schema extension definition

`--owner <owner>`
: The ID of the Azure AD application that is the owner of the schema extension

`-d, --description [description]`
: Description of the schema extension

`-s, --status [status]`
: The lifecycle state of the schema extension. Accepted values are `Available` or `Deprecated`

`-t, --targetTypes [targetTypes]`
: Comma-separated list of Microsoft Graph resource types the schema extension targets

`-p, --properties [properties]`
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

The lifecycle state of the schema extension. The initial state upon creation is `InDevelopment`.
Possible states transitions are from `InDevelopment` to `Available` and `Available` to `Deprecated`.
The target types are the set of Microsoft Graph resource types (that support schema extensions) that this schema extension definition can be applied to. This option is specified as a comma-separated list.
When specifying the JSON string of properties on Windows, you have to escape double quotes in a specific way. Considering the following value for the _properties_ option: `{"Foo":"Bar"}`,
you should specify the value as <code>\`"{""Foo"":""Bar""}"\`</code>.
In addition, when using PowerShell, you should use the `--%` argument.

## Examples

 Update the description of a schema extension

```sh
m365 graph schemaextension set --id MySchemaExtension --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --description "My schema extension" 
```

Update the target types and properties of a schema extension

```sh
m365 graph schemaextension set --id contoso_MySchemaExtension --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --targetTypes "Group,User" --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
```

Update the properties of a schema extension in PowerShell

```PowerShell
graph schemaextension set --id contoso_MySchemaExtension --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties --% \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
```

Change the status of a schema extension to 'Available'

```sh
m365 graph schemaextension set --id contoso_MySchemaExtension --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --status Available
```
