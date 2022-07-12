# planner task get

Retrieve the specified planner task

## Usage

```sh
m365 planner task get [options]
```

## Alias

```sh
m365 planner task details get [options]
```

## Options

`-i, --id  [id]`
: ID of the task. Specify either `id` or `title` but not both. When you specify the task ID, you no longer need to provide the information for `bucket`, `plan`, and `ownerGroup`.

`-t, --title [title]`
: Title of the task. Specify either `id` or `title` but not both.

`--taskId [taskId]`
: (deprecated. Use `id` instead) ID of the task.

`--bucketId [bucketId]`
: ID of the bucket to which the task belongs. Specify `bucketId` or `bucketName` when using `title`.

`--bucketName [bucketName]`
: Name of the bucket to which the task belongs. Specify `bucketId` or `bucketName` when using `title`.

`--planId [planId]`
: ID of the plan to which the task belongs. Specify `planId` or `planTitle` when using `bucketName`.

`--planTitle [planTitle]`
: Title of the plan to which the task belongs. Specify `planId` or `planTitle` when using `bucketName`.

`--planName [planName]`
: (deprecated. Use `planTitle` instead) Title of the plan to which the bucket belongs.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve the specified planner task by id

```sh
m365 planner task get --id "vzCcZoOv-U27PwydxHB8opcADJo-"
```

Retrieve the specified planner task with the title _My Planner Task_ from the bucket named _My Planner Bucket_. Based on the plan with the title _My Planner Plan_ owned by the group _My Planner Group_.

```sh
m365 planner task get --title "My Planner Task" --bucketName "My Planner Bucket" --planTitle "My Planner Plan" --ownerGroupName "My Planner Group"
```
