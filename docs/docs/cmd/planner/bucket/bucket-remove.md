# planner bucket remove

Removes the Microsoft Planner bucket from a plan

## Usage

```sh
m365 planner bucket remove [options]
```

## Options

`-i, --bucketId [bucketId]`
: ID of the bucket to remove. Specify either `bucketId` or `bucketName` but not both.

`-n, --bucketName [bucketName]`
: Name of the bucket to remove. Specify either `bucketId` or `bucketName` but not both.

`--planId [planId]`
: ID of the plan to which the bucket to remove belongs. Specify either `planId` or `planName` when using `bucketName`.

`--planName [planName]`
: Name of the plan to which the bucket to remove belongs. Specify either `planId` or `planName` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--confirm`
: Confirm removal of bucket.


## Examples

Removes the Microsoft Planner bucket rX6L5EVbtUS9nQwVPSo9eMkAGEgM

```sh
m365 planner bucket remove --bucketId "rX6L5EVbtUS9nQwVPSo9eMkAGEgM"
```

Removes the Microsoft Planner bucket My Bucket belonging to plan uO1bj3fdekKuMitpeJqaj8kADBxO

```sh
m365 planner bucket remove --bucketName "My Bucket" --planId "uO1bj3fdekKuMitpeJqaj8kADBxO"
```

Removes the Microsoft Planner bucket My Bucket belonging to plan My Plan in group My Group

```sh
m365 planner bucket remove --bucketName "My Bucket" --planName "My Plan" --ownerGroupName "My Group"
```

## More information

- Delete plannerBucket: [https://docs.microsoft.com/en-us/graph/api/plannerbucket-delete?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/plannerbucket-delete?view=graph-rest-1.0)
