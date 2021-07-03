# planner bucket list

Lists the Microsoft Planner buckets in a plan

## Usage

```sh
m365 planner bucket list [options]
```

## Options

`--planId [planId]`
: ID of the plan to list the buckets of. Specify either `planId` or `planName` but not both.

`--planName [planName]`
: Name of the plan to list the buckets of. Specify either `planId` or `planName` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

--8<-- "docs/cmd/_global.md"

## Examples

Lists the Microsoft Planner buckets in the Plan _xqQg5FS2LkCp935s-FIFm2QAFkHM_

```sh
m365 planner bucket list --planId "xqQg5FS2LkCp935s-FIFm2QAFkHM"
```

Lists the Microsoft Planner buckets in the Plan _My Plan_ owned by group _My Group_

```sh
m365 planner bucket list --planName "My Plan" --ownerGroupName "My Group"
```