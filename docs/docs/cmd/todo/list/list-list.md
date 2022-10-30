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
      },
      {
        "displayName": "My task list",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "none",
        "id": "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
      },
      {
        "displayName": "Flagged Emails",
        "isOwner": true,
        "isShared": false,
        "wellknownListName": "flaggedEmails",
        "id": "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAC98XS0AAAAA=="
      }
    ]
	  ```

=== "Text"

    ```text
    displayName     id                                                                                                      
    --------------  ------------------------------------------------------------------------------------------------------------------------
    Tasks           AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA
    My task list    AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    Flagged Emails  AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAC98XS0AAAAA==
	  ```

=== "CSV"

    ```csv
    displayName,id
    Tasks,AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAAAgESAAAA
    My task list,AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=
    Flagged Emails,AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQAC98XS0AAAAA==
	  ```

