# todo list list

Returns a list of Microsoft To Do task lists

## Usage

```sh
m365 todo list list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get the list of Microsoft To Do task lists

```sh
m365 todo list list
```

## Response

=== "JSON"

    ```json
    [
      {
        "displayName": "Tasks",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "defaultList",
        "id": "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA"
      }
    ]
    ```

=== "Text"

    ```text
    displayName     id                                                                                                      
    --------------  ------------------------------------------------------------------------------------------------------------------------
    Tasks           AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA
    ```

=== "CSV"

    ```csv
    displayName,id
    Tasks,AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA
    ```

=== "Markdown"

    ```md
    # todo list list

    Date: 25/5/2023

    ## Tasks (AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA)
    
    Property | Value
    ---------|-------
    displayName | Tasks
    isOwner | true
    isShared | false
    wellknownListName | defaultList
    id | AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA
    ```
