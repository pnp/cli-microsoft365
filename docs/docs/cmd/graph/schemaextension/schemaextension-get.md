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
      "id": "extuf4mntoy_MySchemaExtension",
      "description": "My Schema Extension",
      "targetTypes": [
        "Group"
      ],
      "status": "InDevelopment",
      "owner": "19b5bd1f-3d5e-404a-80f5-ca840b40a082",
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
    id         : extuf4mntoy_MySchemaExtension
    owner      : 19b5bd1f-3d5e-404a-80f5-ca840b40a082
    properties : [{"name":"myProp1","type":"Integer"},{"name":"myProp2","type":"String"}]
    status     : InDevelopment
    targetTypes: ["Group"]
    ```

=== "CSV"

    ```csv
    id,description,status,owner
    extuf4mntoy_MySchemaExtension,My Schema Extension,InDevelopment,19b5bd1f-3d5e-404a-80f5-ca840b40a082
    ```

=== "Markdown"

    ```md
    # graph schemaextension get --id "extuf4mntoy_MySchemaExtension"

    Date: 2023-05-22

    ## extuf4mntoy_MySchemaExtension

    Property | Value
    ---------|-------
    id | extuf4mntoy\_MySchemaExtension
    description | My Schema Extension
    status | InDevelopment
    owner | 19b5bd1f-3d5e-404a-80f5-ca840b40a082
    ```
