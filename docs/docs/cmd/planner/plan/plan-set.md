# planner plan set

Update title of a specified plan

## Usage

```sh
m365 planner plan set [options]
```

## Options

`-i,--id [id]`
: ID of the plan. Specify either `id` or `title` but not both.

`-t,--title [title]`
: Title of the plan. Specify either `id` or `title` but not both. When `title` is set, specify either `ownerGroupId` or `ownerGroupName`

`--newTitle <newTitle>`
: New title of the plan.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Set new title _MyNewPlan_ for Microsoft Planner plan with id _gndWOTSK60GfPQfiDDj43JgACDCb_

```sh
m365 planner plan set --id "gndWOTSK60GfPQfiDDj43JgACDCb" --newTitle "MyNewPlan"
```

Set new title _MyNewPlan_ for Microsoft Planner plan with original title _MyPlan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan set --title "MyPlan" --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4" --newTitle "MyNewPlan"
```

Set new title _MyNewPlan_ for Microsoft Planner plan with original title _MyPlan_ for Group _My Planner Group_

```sh
m365 planner plan get --title "MyPlan" --ownerGroupName "My Planner Group" --newTitle "MyNewPlan"
```
