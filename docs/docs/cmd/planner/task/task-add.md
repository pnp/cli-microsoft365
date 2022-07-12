# planner task add

Adds a new Microsoft Planner task

## Usage

```sh
m365 planner task add [options]
```

## Options

`-t, --title <title>`
: Title of the task to add.

`--planId [planId]`
: ID of the plan to which the task belongs. Specify either `planId` or `planTitle` but not both.

`--planTitle [planTitle]`
: Title of the plan to which the task belongs. Specify either `planId` or `planTitle` but not both.

`--planName [planName]`
: (deprecated. Use `planTitle` instead) Title of the plan to which the bucket belongs.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--bucketId [bucketId]`
: ID of the bucket to which the task belongs. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--bucketName [bucketName]`
: Name of the bucket to which the task belongs. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--startDateTime [startDateTime]`
: The date and time when the task started. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--dueDateTime [dueDateTime]`
: The date and time when the task is due. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--percentComplete [percentComplete]`
: Percentage of task completion. Number between 0 and 100.

`--assignedToUserIds [assignedToUserIds]`
: The comma-separated IDs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assignedToUserNames [assignedToUserNames]`
: The comma-separated UPNs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assigneePriority [assigneePriority]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--description [description]`
: Description of the task

`--appliedCategories [appliedCategories]`
: Comma-separated categories that should be added to the task. The possible options are: `category1`, `category2`, `category3`, `category4`, `category5` and/or `category6`. Additional info defined [here](https://docs.microsoft.com/graph/api/resources/plannerappliedcategories?view=graph-rest-1.0).

`--previewType [previewType]`
: This sets the type of preview that shows up on the task. The possible values are: `automatic`, `noPreview`, `checklist`, `description`, `reference`. When set to automatic the displayed preview is chosen by the app viewing the task. Default `automatic`.

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--priority [priority]`
: Priority of the task: Urgent, Important, Medium, Low. Or an integer between 0 and 10 (check remarks section for more info). Default value is Medium.

--8<-- "docs/cmd/_global.md"

## Remarks

When you specify the value for `percentComplete`, consider the following:

- when set to 0, the task is considered _Not started_
- when set between 1 and 99, the task is considered _In progress_
- when set to 100, the task is considered _Completed_

When you specify an integer value for `priority`, consider the following:

- values 0 and 1 are interpreted as _Urgent_
- values 2, 3 and 4 are interpreted as _Important_
- values 5, 6 and 7 are interpreted as _Medium_
- values 8, 9 and 10 are interpreted as _Low_

## Examples

Adds a Microsoft Planner task with the name _My Planner Task_ for plan with the ID _8QZEH7b3wkSbGQobscsM5gADCBa_ and for the bucket with the ID _IK8tuFTwQEa5vTonM7ZMRZgAKdna_

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna"
```

Adds a Completed Microsoft Planner task with the name _My Planner Task_ for plan with the title _My Planner Plan_ owned by group _My Planner Group_ and for the bucket with the ID _IK8tuFTwQEa5vTonM7ZMRZgAKdna_

```sh
m365 planner task add --title "My Planner task" --planTitle "My Planner Plan" --ownerGroupName "My Planner Group" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --percentComplete 100
```

Adds a Microsoft Planner task with the name _My Planner Task_ for plan with the ID _8QZEH7b3wkbGQobscsM5gADCBa_ and for the bucket with the ID _IK8tuFTwQEa5vTonM7ZMRZgAKdna_. The new task will be assigned to the users _Allan.Carroll@contoso.com_ and _Ida.Stevens@contoso.com_ and receive a due date for _2021-12-16_

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --assignedToUserNames "Allan.Carroll@contoso.com,Ida.Stevens@contoso.com" --dueDateTime "2021-12-16"
```

Adds a Microsoft Planner task with the name _My Planner Task_ for plan with the ID _8QZEH7b3wkbGQobscsM5gADCBa_ and for the bucket with the ID _IK8tuFTwQEa5vTonM7ZMRZgAKdna_. The new task will be assigned to the users _Allan.Carroll@contoso.com_ and _Ida.Stevens@contoso.com_ who will appear first with the asssignee priority _' !'_ 

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --assignedToUserNames "Allan.Carroll@contoso.com,Ida.Stevens@contoso.com" --asssigneePriority ' !'
```

Adds a Microsoft Planner task with the name _My Planner Task_ for plan with the ID _8QZEH7b3wkbGQobscsM5gADCBa_ and for the bucket with the ID _IK8tuFTwQEa5vTonM7ZMRZgAKdna_. The new task will receive the categories _category1,category3_ and get a preview with the type _noPreview_

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --appliedCategories "category1,category3" --previewType "noPreview"
```
