# planner bucket add

Adds a new Microsoft Planner bucket

## Usage

```sh
m365 planner bucket add [options]
```

## Options

`-n, --name <name>`
: Name of the bucket to add.

`--planId [planId]`
: Plan ID to which the bucket belongs. Specify either `planId` or `planName` but not both.

`--planName [planName]`
: Plan Name to which the bucket belongs. Specify either `planId` or `planName` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

--8<-- "docs/cmd/_global.md"

## Examples

Adds a Microsoft Planner bucket with the name _My Planner Bucket_ for plan with the ID _xqQg5FS2LkCp935s-FIFm2QAFkHM_ with order hint

```sh
m365 planner bucket add --name "My Planner Bucket" --planId "xqQg5FS2LkCp935s-FIFm2QAFkHM" --orderHint " !"
```

Adds a Microsoft Planner bucket with the name _My Planner Bucket_ for plan with the name _My Planner Plan_ owned by group _My Planner Group_

```sh
m365 planner bucket add --name "My Planner Bucket" --planName "My Planner Plan" --ownerGroupName "My Planner Group"
```
