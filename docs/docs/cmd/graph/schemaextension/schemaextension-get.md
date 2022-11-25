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
## Response

=== "JSON"

    ```json
    {
      "id": "extun3653mb_MySchemaExtension",
      "description": "My Schema Extension",
      "targetTypes": [
        "Group"
      ],
      "status": "InDevelopment",
      "owner": "3e789cfc-4c9b-4c5a-a8b0-6b90a28a36f1",
      "properties": [
        {
          "name": "myProp1",
          "type": "Integer"
        },
        {
          "name": "myProp2",
          "type": "String"
        }
      ]
    }
    ```

=== "Text"

    ```text
    description: My Schema Extension
    id         : extun3653mb_MySchemaExtension
    owner      : 3e789cfc-4c9b-4c5a-a8b0-6b90a28a36f1
    properties : [{"name":"myProp1","type":"Integer"},{"name":"myProp2","type":"String"}]
    status     : InDevelopment
    targetTypes: ["Group"]    
    ```

=== "CSV"

    ```csv
    id,description,targetTypes,status,owner,properties
    extun3653mb_MySchemaExtension,My Schema Extension,"[""Group""]",InDevelopment,3e789cfc-4c9b-4c5a-a8b0-6b90a28a36f1,"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]"
    ```
