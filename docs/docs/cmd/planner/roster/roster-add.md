# planner roster add

Creates a new Microsoft Planner Roster

## Usage

```sh
m365 planner roster add [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    The Roster will be automatically deleted when it doesn't contain a plan 24 hours after its creation. Membership information will be completely erased within 30 days of this deletion.

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

!!! important
    To be able to create a new Roster, the Planner Roster creation tenant setting should be enabled. Use the [planner tenant settings list](../tenant/tenant-settings-list.md) command to check if this setting is enabled for your tenant.

## Examples

Creates a new Microsoft Planner Roster

```sh
m365 planner roster add
```

## Response

=== "JSON"

    ```json
    {
      "id": "e6fmvM_yi0OJgvmepE5uj5cAE6qX",
      "assignedSensitivityLabel": null
    }
    ```

=== "Text"

    ```text
    assignedSensitivityLabel: null
    id                      : e6fmvM_yi0OJgvmepE5uj5cAE6qX
    ```

=== "CSV"

    ```csv
    id,assignedSensitivityLabel
    e6fmvM_yi0OJgvmepE5uj5cAE6qX,
    ```
    
## Additional information
Rosters are a new type of container for Microsoft Planner plans. This enables users to create a Planner plan without the need to create a new Microsoft 365 group (with a mailbox, SharePoint site, ...). Access to Roster-contained plans is controlled by the members on the Roster. A Planner Roster can contain only 1 plan.
