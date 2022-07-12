# planner bucket remove

Removes the Microsoft Planner bucket from a plan

## Usage

```sh
m365 planner bucket remove [options]
```

## Options

`-i, --id [id]`
: ID of the bucket to remove. Specify either `id` or `name` but not both.

`-n, --name [name]`
: Name of the bucket to remove. Specify either `id` or `name` but not both.

`--planId [planId]`
: ID of the plan to which the bucket to remove belongs. Specify either `planId` or `planTitle` when using `name`.

`--planTitle [planTitle]`
: Title of the plan to which the bucket to remove belongs. Specify either `planId` or `planTitle` when using `name`.

`--planName [planName]`
: (deprecated. Use `planTitle` instead) Title of the plan to which the bucket belongs.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Removes the Microsoft Planner bucket by ID

```sh
m365 planner bucket remove --id "vncYUXCRBke28qMLB-d4xJcACtNz"
```

Removes the Microsoft Planner bucket by ID without confirmation

```sh
m365 planner bucket remove --id "vncYUXCRBke28qMLB-d4xJcACtNz" --confirm
```

Removes the Microsoft Planner bucket with name _My Bucket_ in the Plan with ID _oUHpnKBFekqfGE_PS6GGUZcAFY7b_

```sh
m365 planner bucket remove --name "My Bucket" --planId "oUHpnKBFekqfGE_PS6GGUZcAFY7b"
```

Removes the Microsoft Planner bucket with name _My Bucket_ in the Plan _My Plan_ owned by group _My Group_

```sh
m365 planner bucket remove --name "My Bucket" --planTitle "My Plan" --ownerGroupName "My Group"
```
