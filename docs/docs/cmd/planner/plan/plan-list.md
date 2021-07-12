# planner plan list

Returns a list of plans associated with a specified group

## Usage

```sh
m365 planner plan list [options]
```

## Options

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Returns a list of Microsoft Planner plans for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan list --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Returns a list of Microsoft Planner plans for Group _My Planner Group_

```sh
m365 planner plan list --ownerGroupName "My Planner Group"
```
