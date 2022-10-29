# todo task get

Gets a specific task from a Microsoft To Do task list

## Usage

```sh
m365 todo task get [options]
```

## Options

`-i, --id <id>`
: The ID of the task in de list.

`--listName [listName]`
: The name of the task list to return tasks from. Specify either `listName` or `listId`, not both

`--listId [listId]`
: The id of the task list to return tasks from. Specify either `listName` or `listId`, not both

--8<-- "docs/cmd/_global.md"

## Examples

Gets a specific task from a Microsoft To Do tasks list based on the name of the list and the task id

```sh
m365 todo task get --listName "My task list" --id "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA="
```

Gets a specific task from a Microsoft To Do tasks list based on the id of the list and the task id

```sh
m365 todo task get --listId "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA==" --id "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA="
```

## Response

=== "JSON"

    ``` json
    {
      "importance": "normal",
      "isReminderOn": false,
      "status": "notStarted",
      "title": "Stay healthy",
      "createdDateTime": "2022-10-23T14:05:09.2673009Z",
      "lastModifiedDateTime": "2022-10-23T14:15:11.3180312Z",
      "hasAttachments": false,
      "categories": [],
      "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
      "body": {
        "content": "",
        "contentType": "text"
      }
    }
    ```

=== "Text"

    ``` text
    createdDateTime     : 2022-10-23T14:05:09.2673009Z
    id                  : AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=
    lastModifiedDateTime: 2022-10-23T14:15:11.3180312Z
    status              : notStarted
    title               : Stay healthy
    ```

=== "CSV"

    ``` text
    id,title,status,createdDateTime,lastModifiedDateTime
    AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=,Stay healthy,notStarted,2022-10-23T14:05:09.2673009Z,2022-10-23T14:15:11.3180312Z
    ```
