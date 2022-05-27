# planner bucket get

Gets the Microsoft Planner bucket in a plan

## Usage

```sh
m365 planner bucket get [options]
```

## Options

`-i, --id [id]`
: ID of the bucket to retrieve details. Specify either `id` or `name` but not both.

`-name, --name [name]`
: Name of the bucket to retrieve details. Specify either `id` or `name` but not both. 

`--planId [planId]`
: Plan ID to which the bucket belongs. Specify either `planId` or `planName` when using `name`.

`--planName [planName]`
: Plan Name to which the bucket belongs. Specify either `planId` or `planName` when using `name`.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

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
m365 planner bucket get --name "Planner Bucket A" --planName "My Plan" --ownerGroupName "My Group"
```

Gets the Microsoft Planner bucket in the Plan _My Plan_ owned by groupId ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e

```sh
m365 planner bucket get --name "Planner Bucket A" --planName "My Plan" --ownerGroupId "ee0f40fc-b2f7-45c7-b62d-11b90dd2ea8e"
```