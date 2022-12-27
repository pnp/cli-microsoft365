# planner plan list

Returns a list of plans associated with a specified group

## Usage

```sh
m365 planner plan list [options]
```

## Options

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Returns a list of Microsoft Planner plans for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan list --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Returns a list of Microsoft Planner plans for Group _My Planner Group_

```sh
m365 planner plan list --ownerGroupName "My Planner Group"
```

## Response

=== "JSON"

    ```json
    [
      {
        "createdDateTime": "2015-03-30T18:36:49.2407981Z",
        "owner": "ebf3b108-5234-4e22-b93d-656d7dae5874",
        "title": "My Planner Plan",
        "id": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "95e27074-6c4a-447a-aa24-9d718a0b86fa"
          },
          "application": {
            "displayName": null,
            "id": "ebf3b108-5234-4e22-b93d-656d7dae5874"
          }
        },
        "container": {
          "containerId": "ebf3b108-5234-4e22-b93d-656d7dae5874",
          "type": "group",
          "url": "https://graph.microsoft.com/v1.0/groups/ebf3b108-5234-4e22-b93d-656d7dae5874"
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                            title            createdDateTime               owner
    ----------------------------  ---------------  ----------------------------  ------------------------------------
    xqQg5FS2LkCp935s-FIFm2QAFkHM  My Planner Plan  2015-03-30T18:36:49.2407981Z  ebf3b108-5234-4e22-b93d-656d7dae5874
    ```

=== "CSV"

    ```csv
    id,title,createdDateTime,owner
    xqQg5FS2LkCp935s-FIFm2QAFkHM,My Planner Plan,2015-03-30T18:36:49.2407981Z,ebf3b108-5234-4e22-b93d-656d7dae5874
    ```

=== "Markdown"

    ```md
    # planner plan list --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"

    Date: 27/12/2022

    ## My Planner Plan (xqQg5FS2LkCp935s-FIFm2QAFkHM)

    Property | Value
    ---------|-------
    createdDateTime | 2021-12-12T14:00:50.4522129Z
    owner | 233e43d0-dc6a-482e-9b4e-0de7a7bce9b4
    title | My Planner Plan
    id | xqQg5FS2LkCp935s-FIFm2QAFkHM
    createdBy | {"user":{"displayName":null,"id":"dd8b99a7-77c6-4238-a609-396d27844921"},"application":{"displayName":null,"id":"09abbdfd-ed23-44ee-a2d9-a627aa1c90f3"}}
    container | {"containerId":"233e43d0-dc6a-482e-9b4e-0de7a7bce9b4","type":"group","url":"https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"}
    ```
