# planner plan remove

Removes the Microsoft Planner plan

## Usage

```sh
m365 planner plan remove [options]
```

## Options

`i, --id [id]`
: ID of the plan to remove. Specify either `id` or `title` but not both.

`-t, --title [title]`
: Title of the plan to remove. Specify either `id` or `title` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes the Microsoft Planner plan by ID

```sh
m365 planner plan remove --id gndWOTSK60GfPQfiDDj43JgACDCb
```

Removes the Microsoft Planner plan with title _My Plan_ in group with specific ID

```sh
m365 planner plan remove --title "My Plan" --ownerGroupId 00000000-0000-0000-0000-000000000000
```

Removes the Microsoft Planner plan with title _My Plan_ in group with name _My Planner Group_ without confirmation prompt

```sh
m365 planner plan remove --title "My Plan" --ownerGroupName "My Planner Group" --confirm
```
