# planner task get

Retrieve the the specified planner task

## Usage

```sh
m365 planner task get [options]
```

## Options

`-i, --id <id>`
: ID of the task to retrieve details from

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command uses an API that is currently in preview to enrich the results with the `priority` field. Keep in mind that this preview API is subject to change once the API reached general availability.

## Examples

Retrieve the the specified planner task

```sh
m365 planner task get --id 'vzCcZoOv-U27PwydxHB8opcADJo-'
```
