# pp dataverse table get

List a dataverse table in a given environment

## Usage

```sh
pp dataverse table get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment to list a table for.

`-n, --name<name>`
: The name of the dataverse table to retrieve rows from.

`-a, --asAdmin`
: Set, to retrieve the dataverse table as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Examples

List a table for the given environment

```sh
m365 pp dataverse table get -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "aaduser"
```

List a table for the given environment as Admin

```sh
m365 pp dataverse table get -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "aaduser" --asAdmin
```

## Response

=== "JSON"

    ```json
    {
      "MetadataId": "84f4c125-474d-ed11-bba1-000d3a2caf7f",
      "IsCustomEntity": true,
      "IsManaged": false,
      "SchemaName": "aaduser",
      "IconVectorName": null,
      "LogicalName": "aaduser",
      "EntitySetName": "aadusers",
      "IsActivity": false,
      "DataProviderId": null,
      "IsRenameable": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "isrenameable"
      },
      "IsCustomizable": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "iscustomizable"
      },
      "CanCreateForms": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "cancreateforms"
      },
      "CanCreateViews": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "cancreateviews"
      },
      "CanCreateCharts": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "cancreatecharts"
      },
      "CanCreateAttributes": {
        "Value": true,
        "CanBeChanged": false,
        "ManagedPropertyLogicalName": "cancreateattributes"
      },
      "CanChangeTrackingBeEnabled": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "canchangetrackingbeenabled"
      },
      "CanModifyAdditionalSettings": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "canmodifyadditionalsettings"
      },
      "CanChangeHierarchicalRelationship": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "canchangehierarchicalrelationship"
      },
      "CanEnableSyncToExternalSearchIndex": {
        "Value": true,
        "CanBeChanged": true,
        "ManagedPropertyLogicalName": "canenablesynctoexternalsearchindex"
      }
    }
    ```

=== "Text"

    ```text
    EntitySetName: aadusers
    IsManaged    : false
    LogicalName  : aaduser
    SchemaName   : aaduser
    ```

=== "CSV"

    ```csv
    SchemaName,EntitySetName,LogicalName,IsManaged
    aaduser,aadusers,aaduser,
    ```
