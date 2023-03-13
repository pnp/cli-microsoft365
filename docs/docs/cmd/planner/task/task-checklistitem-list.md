# planner task checklistitem list

Lists the checklist items of a Planner task.

## Usage

```sh
m365 planner task checklistitem list [options]
```

## Options

`-i, --taskId <taskId>`
: ID of the task

--8<-- "docs/cmd/_global.md"

## Examples

Lists the checklist items of a Planner task.

```sh
m365 planner task checklistitem list --taskId 'vzCcZoOv-U27PwydxHB8opcADJo-'
```

## Response

=== "JSON"

    ```json
    {
      "4e3c8841-560c-436e-ba06-cc7731680d59": {
        "isChecked": false,
        "title": "Communicate with customer",
        "orderHint": "8585269209601773376",
        "lastModifiedDateTime": "2023-01-25T22:32:05.3002431Z",
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
    id                                    title                      isChecked
    ------------------------------------  -------------------------  ---------
    4e3c8841-560c-436e-ba06-cc7731680d59  Communicate with customer  false
    ```

=== "CSV"

    ```csv
    id,title,isChecked
    4e3c8841-560c-436e-ba06-cc7731680d59,Communicate with customer,
    ```

=== "Markdown"

    ```md
    # planner task checklistitem list --taskId "OopX1ANphEu7Lm4-0tVtl5cAFRGQ"

    Date: 25/1/2023

    ## Communicate with customer (4e3c8841-560c-436e-ba06-cc7731680d59)

    Property | Value
    ---------|-------
    id | 4e3c8841-560c-436e-ba06-cc7731680d59
    isChecked | false
    title | Communicate with customer
    orderHint | 8585269209601773376
    lastModifiedDateTime | 2023-01-25T22:32:05.3002431Z
    lastModifiedBy | {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"}}
    ```
