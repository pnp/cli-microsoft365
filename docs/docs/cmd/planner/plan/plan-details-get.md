# planner plan details get

Retrieve the planner details about the specified plan.

## Usage

```sh
m365 planner plan details get [options]
```

## Options

`-i, --planId [planId]`
: ID of the plan. Specify either `planId` or `planTitle` but not both.

`-t, --planTitle [planTitle]`
: Title of the plan. Specify either `planId` or `planTitle` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Returns the Microsoft Planner plan details with id _gndWOTSK60GfPQfiDDj43JgACDCb_

```sh
m365 planner plan details get --planId "gndWOTSK60GfPQfiDDj43JgACDCb"
```

Returns the Microsoft Planner plan details with title _MyPlan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan details get --planTitle "MyPlan" --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Returns the Microsoft Planner plan details with title _MyPlan_ for Group _My Planner Group_

```sh
m365 planner plan details get --planTitle "MyPlan" --ownerGroupName "My Planner Group"
```
