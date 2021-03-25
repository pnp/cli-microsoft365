# planner plan add

Adds a new Microsoft Planner plan

## Usage

```sh
m365 planner plan add [options]
```

## Options

`-t, --title <title>`
: Title of the plan to add.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. A valid group must exist before this option can be set. Specify either `ownerGroupId` or `ownerGroupName` but not both.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. A valid group must exist before this option can be set. Specify either `ownerGroupId` or `ownerGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Adds a Microsoft Planner plan with the name _My Planner Plan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan add --title "My Planner Plan" --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Adds a Microsoft Planner plan with the name _My Planner Plan_ for Group _My Planner Group_

```sh
m365 planner plan add --title "My Planner Plan" --ownerGroupName "My Planner Group"
```
