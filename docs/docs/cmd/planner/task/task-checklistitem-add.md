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
      "b65db83e-d777-49db-8cdd-57a41b86f48c": {
        "isChecked": false,
        "title": "Communicate with customer",
        "orderHint": "8585269221832287402",
        "lastModifiedDateTime": "2023-01-25T22:11:42.2488405Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "b2091e18-7882-4efe-b7d1-90703f5a5c65"
          }
        }
      }
    }
    ```

=== "Text"

    ```txt
    id       : b65db83e-d777-49db-8cdd-57a41b86f48c
    isChecked: false
    title    : Communicate with customer
    ```

=== "CSV"

    ```csv
    id,title,isChecked
    b65db83e-d777-49db-8cdd-57a41b86f48c,Communicate with customer,
    ```

=== "Markdown"

    ```md
    # planner task checklistitem add --taskId "OopX1ANphEu7Lm4-0tVtl5cAFRGQ" --title "Communicate with customer"

    Date: 25/1/2023

    ## Communicate with customer (b65db83e-d777-49db-8cdd-57a41b86f48c)

    Property | Value
    ---------|-------
    id | b65db83e-d777-49db-8cdd-57a41b86f48c
    isChecked | false
    title | Communicate with customer
    orderHint | 8585269209601773376
    lastModifiedDateTime | 2023-01-25T22:32:05.3002431Z
    lastModifiedBy | {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"}}
    ```
