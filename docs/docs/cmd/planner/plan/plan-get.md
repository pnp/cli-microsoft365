# planner plan get

Retrieve information about the specified plan

## Usage

```sh
m365 planner plan get [options]
```

## Alias

```sh
m365 planner plan details get [options]
```

## Options

`-i, --id [id]`
: ID of the plan. Specify either `id` or `title` but not both.

`-t, --title [title]`
: Title of the plan. Specify either `id` or `title` but not both.

`--planId [planId]`
: (deprecated. Use `id` instead) ID of the plan. Specify either `planId` or `planTitle` but not both.

`---planTitle [planTitle]`
: (deprecated. Use `title` instead) Title of the plan. Specify either `planId` or `planTitle` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title` or the deprecated `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title` or the deprecated `planTitle`.

--8<-- "docs/cmd/_global.md"

## Examples

Returns the Microsoft Planner plan with id _gndWOTSK60GfPQfiDDj43JgACDCb_

```sh
m365 planner plan get --id "gndWOTSK60GfPQfiDDj43JgACDCb"
```

Returns the Microsoft Planner plan with title _MyPlan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan get --title "MyPlan" --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Returns the Microsoft Planner plan with title _MyPlan_ for Group _My Planner Group_

```sh
m365 planner plan get --title "MyPlan" --ownerGroupName "My Planner Group"
```
