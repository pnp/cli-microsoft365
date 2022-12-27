# planner bucket get

Gets the Microsoft Planner bucket in a plan

## Usage

```sh
m365 planner bucket get [options]
```

## Options

`-i, --id [id]`
: ID of the bucket to retrieve details. Specify either `id` or `name` but not both.

`-n, --name [name]`
: Name of the bucket to retrieve details. Specify either `id` or `name` but not both. 

`--planId [planId]`
: ID of the plan to which the bucket belongs. Specify either `planId` or `planTitle` when using `name`.

`--planTitle [planTitle]`
: Title of the plan to which the bucket belongs. Specify either `planId` or `planTitle` when using `name`.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the specified Microsoft Planner bucket 

```sh
m365 planner bucket get --id "5h1uuYFk4kKQ0hfoTUkRLpgALtYi"
```

Gets the Microsoft Planner bucket in the PlanId xqQg5FS2LkCp935s-FIFm2QAFkHM

```sh
m365 planner bucket get --name "Planner Bucket A" --planId "xqQg5FS2LkCp935s-FIFm2QAFkHM"
```

Gets the Microsoft Planner bucket in the Plan _My Plan_ owned by group _My Group_

```sh
m365 planner bucket get --name "Planner Bucket A" --planTitle "My Plan" --ownerGroupName "My Group"
```

Gets the Microsoft Planner bucket in the Plan _My Plan_ owned by groupId ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e

```sh
m365 planner bucket get --name "Planner Bucket A" --planTitle "My Plan" --ownerGroupId "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e"
```

## Response

=== "JSON"

    ```json
    {
      "name": "My Planner Bucket",
      "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
      "orderHint": "8585363889524958496",
      "id": "ttEB_Uj690STdR3GC1MIDZgANq1U"
    }
    ```

=== "Text"

    ```text
    id       : ttEB_Uj690STdR3GC1MIDZgANq1U
    name     : My Planner Bucket
    orderHint: 8585363889524958496
    planId   : xqQg5FS2LkCp935s-FIFm2QAFkHM
    ```

=== "CSV"

    ```csv
    id,name,planId,orderHint
    ttEB_Uj690STdR3GC1MIDZgANq1U,My Planner Bucket,xqQg5FS2LkCp935s-FIFm2QAFkHM,8585363889524958496
    ```

=== "Markdown"

    ```md
    # planner bucket get --id 'ttEB_Uj690STdR3GC1MIDZgANq1U"

    Date: 27/12/2022

    ## My Planner Bucket (ttEB_Uj690STdR3GC1MIDZgANq1U)

    Property | Value
    ---------|-------
    name | My Planner Bucket
    planId | xqQg5FS2LkCp935s-FIFm2QAFkHM
    orderHint | 8585363889524958496
    id | ttEB_Uj690STdR3GC1MIDZgANq1U
    ```
