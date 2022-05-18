# planner task checklistitem add

Adds a new checklist item to a Planner task

## Usage

```sh
m365 planner task checklistitem add [options]
```

## Options

`-i, --taskId <taskId>`
: ID of the task.

`-t, --title <title>`
: Title of the checklist item.

`--isChecked`
: Mark the checklist item as checked.

--8<-- "docs/cmd/_global.md"

## Examples

Adds an unchecked checklist item with title _My checklist item_ to a Microsoft Planner task with ID _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task checklistitem add --taskId 2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2 --title "My checklist item"
```

Adds a checked checklist item with title _My checklist item_ to a Microsoft Planner task with ID _2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2_

```sh
m365 planner task checklistitem add --taskId 2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2 --title "My checklist item" --isChecked
```
