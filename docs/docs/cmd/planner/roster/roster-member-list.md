# planner roster member list

Lists members of the specified Microsoft Planner Roster

## Usage

```sh
m365 planner roster member list [options]
```

## Options

`--rosterId <rosterId>`
: ID of the Planner Roster.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Lists members of the specified Microsoft Planner Roster

```sh
m365 planner roster member list --rosterId tYqYlNd6eECmsNhN_fcq85cAGAnd
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "b3a1be03-54a5-43d2-b4fb-6562fe9bec0b",
        "userId": "2056d2f6-3257-4253-8cfc-b73393e414e5",
        "tenantId": "5b7b813c-2339-48cd-8c51-bd4fcb269420",
        "roles": []
      }
    ]
    ```

=== "Text"

    ```text
    id      : b3a1be03-54a5-43d2-b4fb-6562fe9bec0b
    roles   : []
    tenantId: 5b7b813c-2339-48cd-8c51-bd4fcb269420
    userId  : 2056d2f6-3257-4253-8cfc-b73393e414e5
    ```

=== "CSV"

    ```csv
    id,userId,tenantId,roles
    b3a1be03-54a5-43d2-b4fb-6562fe9bec0b,2056d2f6-3257-4253-8cfc-b73393e414e5,5b7b813c-2339-48cd-8c51-bd4fcb269420,[]
    ```
