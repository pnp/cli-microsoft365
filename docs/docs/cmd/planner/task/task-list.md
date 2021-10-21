# planner task list

Lists planner tasks in a bucket, plan, or tasks for the currently logged in user

## Usage

```sh
m365 planner task list [options]
```

## Options

`--bucketId [bucketId]`
: ID of the bucket to list the tasks of. To retrieve tasks from a bucket, specify `bucketId` or `bucketName`, but not both.

`--bucketName [bucketName]`
: Name of the bucket to list the tasks of. To retrieve tasks from a bucket, specify `bucketId` or `bucketName`, but not both.

`--planId [planId]`
: ID of the plan to list the tasks of. Specify `planId` or `planName` when using `bucketName`.

`--planName [planName]`
: Name of the plan to list the tasks of. Specify `planId` or `planName` when using `bucketName`.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

--8<-- "docs/cmd/_global.md"

## Examples

List tasks for the currently logged in user

```sh
m365 planner task list
```

List the Microsoft Planner tasks in the plan _iVPMIgdku0uFlou-KLNg6MkAE1O2_

```sh
m365 planner task list --planId "iVPMIgdku0uFlou-KLNg6MkAE1O2"`
```

List the Microsoft Planner tasks in the plan _My Plan_ in group _My Group_

```sh
m365 planner task list --planName "My Plan" --ownerGroupName "My Group"
```

List the Microsoft Planner tasks in the bucket _FtzysDykv0-9s9toWiZhdskAD67z_

```sh
m365 planner task list --bucketId "FtzysDykv0-9s9toWiZhdskAD67z"
```

List the Microsoft Planner tasks in the bucket _My Bucket_ belonging to plan _iVPMIgdku0uFlou-KLNg6MkAE1O2_

```sh
m365 planner task list --bucketName "My Bucket" --planId "iVPMIgdku0uFlou-KLNg6MkAE1O2"
```

List the Microsoft Planner tasks in the bucket _My Bucket_ belonging to plan _My Plan_ in group _My Group_

```sh
m365 planner bucket tasks list --bucketName "My Bucket" --planName "My Plan" --ownerGroupName "My Group"
```
