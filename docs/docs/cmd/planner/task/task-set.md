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
: Title of the task to add.

`--bucketId [bucketId]`
: Bucket ID to which the task belongs. Specify either `bucketId` or `bucketName` but not both.

`--bucketName [bucketName]`
: Bucket Name to which the task belongs. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--planId [planId]`
: Plan ID to which the task belongs. Specify either `planId` or `planName` but not both.

`--planName [planName]`
: Plan Name to which the task belongs. Specify either `planId` or `planName` but not both.

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
  - When set to 0, the task is considered _Not started_. 
  - When set between 1 and 99, the task is considered _In progress_.
  - When set to 100, the task is considered _Completed_.

`--assignedToUserIds [assignedToUserIds]`
: The comma-separated IDs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assignedToUserNames [assignedToUserNames]`
: The comma-separated UPNs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--description [description]`
: Description of the task

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--assigneePriority [assigneePriority]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--conversationThreadId [conversationThreadId]`
: Thread id of the conversation on the task. This is the id of the conversation thread object created in the group.

`--appliedCategories [appliedCategories]`
: The comma-separated categories to which the task has been applied. Possible values see [here](https://docs.microsoft.com/en-us/graph/api/resources/plannerappliedcategories?view=graph-rest-1.0). Values are set to true. `category1,category3`

--8<-- "docs/cmd/_global.md"

## Examples

Updates a Microsoft Planner task with the name _My Planner Task_ for task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_

```sh
m365 planner task set -i "Z-RLQGfppU6H3663DBzfs5gAMD3o" --title "My Planner Task"
```

Updates a Microsoft Planner task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_ to the bucket named _My Planner Bucket_. Based on the plan with the name _My Planner Plan_ owned by the group _My Planner Group_

```sh
m365 planner task set  -i "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --bucketName "My Planner Bucket" --planName "My Planner Plan" --ownerGroupName "My Planner Group"
```

Updates a Microsoft Planner task with the ID _Z-RLQGfppU6H3663DBzfs5gAMD3o_ to complete and adds _Category1_ and _Category3_.

```sh
m365 planner task set -i "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2"  --percentComplete 50 --appliedCategories "category1,category3"
```

