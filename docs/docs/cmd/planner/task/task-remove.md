# planner task remove

Removes the Microsoft Planner task from a plan

## Usage

```sh
m365 planner task remove [options]
```

## Options

`--id [id]`
: ID of the task to remove. Specify either `id` or `title` but not both.

`--title [title]`
: Title of the task to remove. Specify either `id` or `title` but not both.

`--bucketId [bucketId]`
: ID of the bucket to which the task to remove belongs. Specify either `bucketId` or `bucketName` but not both.

`--bucketName [bucketName]`
: Name of the bucket to which the task to remove belongs. Specify either `bucketId` or `bucketName` but not both.

`--planId [planId]`
: ID of the plan to which the task to remove belongs. Specify either `planId` or `planTitle` when using `title`.

`--planTitle [planTitle]`
: Title of the plan to which the task to remove belongs. Specify either `planId` or `planTitle` when using `title`.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Removes the Microsoft Planner task by ID

```sh
m365 planner task remove --id "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2"
```

Removes the Microsoft Planner task by ID without confirmation

```sh
m365 planner task remove --id "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --confirm
```

Removes the Microsoft Planner task with title _My Task_ in the bucket with ID _vncYUXCRBke28qMLB-d4xJcACtNz_

```sh
m365 planner task remove --title "My Task" --bucketId "vncYUXCRBke28qMLB-d4xJcACtNz" 
```

Removes the Microsoft Planner task with title _My Task_ in the bucket with name _My Bucket_ in the Plan with ID _oUHpnKBFekqfGE_PS6GGUZcAFY7b_

```sh
m365 planner task remove --title "My Task" --bucketName "My Bucket" --planId "oUHpnKBFekqfGE_PS6GGUZcAFY7b"
```

Removes the Microsoft Planner task with title _My Task_ in the bucket with name _My Bucket_ in the Plan _My Plan_ owned by group _My Group_

```sh
m365 planner task remove --title "My Task" --bucketName "My Bucket" --planTitle "My Plan" --ownerGroupName "My Group"
```

Removes the Microsoft Planner task with title _My Task_ in the bucket with name _My Bucket_ in the Plan _My Plan_ owned by group with ID _00000000-0000-0000-0000-000000000000_

```sh
m365 planner task remove --title "My Task" --bucketName "My Bucket" --planTitle "My Plan" --ownerGroupId "00000000-0000-0000-0000-000000000000"
```
