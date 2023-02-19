# planner roster get

Gets information about the specific Microsoft Planner Roster.

## Usage

```sh
m365 planner roster get [options]
```

## Options

`--id <id>`
: ID of the Planner Roster.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Gets information about a specific Planner Roster.

```sh
m365 planner roster get --id tYqYlNd6eECmsNhN_fcq85cAGAnd
```

## Response

=== "JSON"

    ```json
    {
      "id": "tYqYlNd6eECmsNhN_fcq85cAGAnd",
      "assignedSensitivityLabel": null
    }
    ```

=== "Text"

    ```text
    assignedSensitivityLabel: null
    id                      : tYqYlNd6eECmsNhN_fcq85cAGAnd
    ```

=== "CSV"

    ```csv
    id,assignedSensitivityLabel
    tYqYlNd6eECmsNhN_fcq85cAGAnd,
    ```

=== "Markdown"

    ```md
    # planner roster get --id "tYqYlNd6eECmsNhN_fcq85cAGAnd"

    Date: 1/30/2023

    ## undefined (tYqYlNd6eECmsNhN_fcq85cAGAnd)

    Property | Value
    ---------|-------
    id | tYqYlNd6eECmsNhN_fcq85cAGAnd
    assignedSensitivityLabel | null
    ```

## Additional information

Rosters are a new type of container for Microsoft Planner plans. This enables users to create a Planner plan without the need to create a new Microsoft 365 group (with a mailbox, SharePoint site, ...). Access to Roster-contained plans is controlled by the members on the Roster. A Planner Roster can contain only 1 plan.
