# todo task list

List tasks from a Microsoft To Do task list

## Usage

```sh
m365 todo task list [options]
```

## Options

`--listName [listName]`
: The name of the task list to return tasks from. Specify either `listName` or `listId`, not both

`--listId [listId]`
: The id of the task list to return tasks from. Specify either `listName` or `listId`, not both

--8<-- "docs/cmd/_global.md"

## Examples

List tasks from Microsoft To Do tasks list with the name _My task list_

```sh
m365 todo task list --listName "My task list"
```

List tasks from Microsoft To Do tasks list with the id AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA==

```sh
m365 todo task list --listId "AQMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MAAuAAADMN-7V4K8g0q_adetip1DygEAxMBBaLl1lk_dAn8KkjfXKQABF-BAgwAAAA=="
```

## Response

=== "JSON"

    ```json
    [
      {
        "importance": "high",
        "isReminderOn": true,
        "status": "notStarted",
        "title": "New task",
        "createdDateTime": "2022-10-29T11:03:20.9175176Z",
        "lastModifiedDateTime": "2022-10-29T11:13:23.6672968Z",
        "hasAttachments": false,
        "categories": [],
        "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=",
        "body": {
          "content": "I should not forget this",
          "contentType": "text"
        },
        "dueDateTime": {
          "dateTime": "2023-01-01T00:00:00.0000000",
          "timeZone": "UTC"
        },
        "reminderDateTime": {
          "dateTime": "2023-01-01T12:00:00.0000000",
          "timeZone": "UTC"
        }
      },
      {
        "importance": "normal",
        "isReminderOn": false,
        "status": "notStarted",
        "title": "Important task",
        "createdDateTime": "2022-10-29T10:55:31.1510823Z",
        "lastModifiedDateTime": "2022-10-29T11:05:33.5899799Z",
        "hasAttachments": false,
        "categories": [],
        "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtPAAA=",
        "body": {
          "content": "",
          "contentType": "text"
        }
      },
      {
        "importance": "normal",
        "isReminderOn": false,
        "status": "notStarted",
        "title": "Another task",
        "createdDateTime": "2022-10-29T10:54:49.6380703Z",
        "lastModifiedDateTime": "2022-10-29T11:04:51.1161127Z",
        "hasAttachments": false,
        "categories": [],
        "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtOAAA=",
        "body": {
          "content": "",
          "contentType": "text"
        }
      }
    ]
	  ```

=== "Text"

    ```text
    id                                                                                                                                                        title           status      createdDateTime               lastModifiedDateTime
  --------------------------------------------------------------------------------------------------------------------------------------------------------  --------------  ----------  ----------------------------  ----------------------------
  AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=  New task        notStarted  2022-10-29T11:03:20.9175176Z  2022-10-29T11:13:23.6672968Z
  AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtPAAA=  Important task  notStarted  2022-10-29T10:55:31.1510823Z  2022-10-29T11:05:33.5899799Z
  AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtOAAA=  Another task    notStarted  2022-10-29T10:54:49.6380703Z  2022-10-29T11:04:51.1161127Z
	  ```

=== "CSV"

    ```csv
    id,title,status,createdDateTime,lastModifiedDateTime
    AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtSAAA=,New task,notStarted,2022-10-29T11:03:20.9175176Z,2022-10-29T11:13:23.6672968Z
    AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtPAAA=,Important task,notStarted,2022-10-29T10:55:31.1510823Z,2022-10-29T11:05:33.5899799Z
    AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAL3xdLTAADEwEFouXWWT50CfwqSN9cpAAL3xhtOAAA=,Another task,notStarted,2022-10-29T10:54:49.6380703Z,2022-10-29T11:04:51.1161127Z
	  ```
