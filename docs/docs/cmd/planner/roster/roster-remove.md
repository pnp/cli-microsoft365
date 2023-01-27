# planner roster remove

Removes a Microsoft Planner Roster

## Usage

```sh
m365 planner roster remove [options]
```

## Options

`i, --id <id>`
: ID of the Planner Roster.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks
!!! attention
    Deleting a Planner Roster will also delete the plan within the Roster.

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Removes a Planner Roster

```sh
m365 planner roster remove --id tYqYlNd6eECmsNhN_fcq85cAGAnd
```

Removes a Planner Roster without confirmation prompt

```sh
m365 planner roster remove --id tYqYlNd6eECmsNhN_fcq85cAGAnd --confirm
```

## Response

The command won't return a response on success.

## Additional information
- Rosters are a new type of container for Microsoft Planner plans. This enables users to create a Planner plan without the need to create a new Microsoft 365 group (with a mailbox, SharePoint site, ...). Access to Roster-contained plans is controlled by the members on the Roster. A Planner Roster can contain only 1 plan.
