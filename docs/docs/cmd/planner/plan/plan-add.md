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

`--shareWithUserIds [shareWithUserIds]`
: The comma-separated IDs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

`--shareWithUserNames [shareWithUserNames]`
: The comma-separated UPNs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

Related to the options `--shareWithUserIds` and `--shareWithUserNames`. If you are leveraging Microsoft 365 groups, use the `aad o365group user` commands to manage group membership to share the [group's](https://pnp.github.io/cli-microsoft365/cmd/aad/o365group/o365group-user-add/) plan. You can also add existing members of the group to this collection though it is not required for them to access the plan owned by the group.

## Examples

Adds a Microsoft Planner plan with the name _My Planner Plan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupId '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
```

Adds a Microsoft Planner plan with the name _My Planner Plan_ for Group _My Planner Group_

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupName 'My Planner Group'
```

Adds a Microsoft Planner plan with the name _My Planner Plan_ for Group _My Planner Group_ and share it with the users _Allan.Carroll@contoso.com_ and _Ida.Stevens@contoso.com_

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupName 'My Planner Group' --shareWithUserNames 'Allan.Carroll@contoso.com,Ida.Stevens@contoso.com'
```
