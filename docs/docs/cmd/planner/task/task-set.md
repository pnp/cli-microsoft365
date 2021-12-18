# planner task set

Updates a Microsoft Planner task

## Usage

```sh
m365 planner task set [options]
```

## Options

`-i, --id <id>`
: ID of the task.

`-t, --title [title]`
: New title of the task.

`--bucketId [bucketId]`
: ID of the bucket to move the task to. Specify either `bucketId` or `bucketName` but not both.

`--bucketName [bucketName]`
: Name of the bucket to move the task to. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--planId [planId]`
: ID of the plan to move the task to. Specify either `planId` or `planName` but not both.

`--planName [planName]`
: Name of the plan to move the task to. Specify either `planId` or `planName` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planName`.

`--startDateTime [startDateTime]`
: The date and time when the task started. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--dueDateTime [dueDateTime]`
: The date and time when the task is due. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--percentComplete [percentComplete]`
: Percentage of task completion. Number between 0 and 100.

`--assignedToUserIds [assignedToUserIds]`
: Comma-separated IDs of the assignees that should be added to the task assignment. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assignedToUserNames [assignedToUserNames]`
: Comma-separated UPNs of the assignees that should be added to the task assignment. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--description [description]`
: Description of the task

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view

`--assigneePriority [assigneePriority]`
: Hint used to order items of this type in a list view

`--appliedCategories [appliedCategories]`
: Comma-separated categories that should be added to the task

--8<-- "docs/cmd/_global.md"

## Remarks

When you specify the value for `percentageComplete`, consider the following:

- when set to 0, the task is considered _Not started_
- when set between 1 and 99, the task is considered _In progress_
- when set to 100, the task is considered _Completed_

You can add up to 6 categories to the task. An example to add _category1_ and _category3_ would be `category1,category3`.

## Examples

Updates a Microsoft Planner task name to _My Planner Task_ for the task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_

```sh
m365 planner task set --id "Z-RLQGfppU6H3663DBzfs5gAMD3o" --title "My Planner Task"
```

Moves a Microsoft Planner task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_ to the bucket named _My Planner Bucket_. Based on the plan with the name _My Planner Plan_ owned by the group _My Planner Group_

```sh
m365 planner task set  --id "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --bucketName "My Planner Bucket" --planName "My Planner Plan" --ownerGroupName "My Planner Group"
```

Marks a Microsoft Planner task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_ as 50% complete and assigned to categories 1 and 3.

```sh
m365 planner task set --id "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2"  --percentComplete 50 --appliedCategories "category1,category3"
```

## Additional information

- Using order hints in Planner: [https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0](https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0)
- Applied categories in Planner: [https://docs.microsoft.com/graph/api/resources/plannerappliedcategories?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/plannerappliedcategories?view=graph-rest-1.0)
