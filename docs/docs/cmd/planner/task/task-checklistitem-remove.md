# planner task checklistitem remove

Removes the checklist item from the Planner task.

## Usage

```sh
m365 planner task checklistitem remove [options]
```

## Options

`-i, --id <id>`
: ID of the checklist item.

`--taskId <taskId>`
: ID of the task.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Removes a checklist item with the id _40012_ from the Planner task with the id _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task checklistitem remove --id "40012" --taskId "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" 
```

Removes a checklist item with the id _40012_ from the Planner task with the id _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_ without prompt

```sh
m365 planner task checklistitem remove --id "40012" --taskId "2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2" --confirm
```
