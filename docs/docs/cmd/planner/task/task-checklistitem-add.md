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

## Response

=== "JSON"

    ```json
    {
      "isChecked": false,
      "title": "My checklist item",
      "orderHint": "8585267801845015741",
      "lastModifiedDateTime": "2023-01-27T13:38:20.9760066Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "fe36f75e-c103-410b-a18a-2bf6df06ac3a"
        }
      },
      "id": "c6dd3b81-75ea-4bbe-af28-c4c4f5b2f7c5"
    }
    ```

=== "Text"

    ```text
    id       : c6dd3b81-75ea-4bbe-af28-c4c4f5b2f7c5
    isChecked: false
    title    : My checklist item
    ```

=== "CSV"

    ```csv
    c6dd3b81-75ea-4bbe-af28-c4c4f5b2f7c5,My checklist item,
    ```

=== "Markdown"

    ```md
    # planner task checklistitem add --taskId "fiinRHmjFk2Yc5n58hNC1JcABptv" --title "My checklist item"

    Date: 27/01/2023

    ## My checklist item (30941772-2949-4d49-aad1-d3f6ee42b265)

    Property | Value
    ---------|-------
    isChecked | false
    title | My checklist item
    orderHint | 8585267800420526438
    lastModifiedDateTime | 2023-01-27T13:40:43.4249369Z
    lastModifiedBy | {"user":{"displayName":null,"id":"fe36f75e-c103-410b-a18a-2bf6df06ac3a"}}
    id | 30941772-2949-4d49-aad1-d3f6ee42b265
    ```

