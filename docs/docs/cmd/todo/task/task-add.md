# todo task add

Adds a task to a Microsoft To Do list

## Usage

```sh
m365 todo task add [options]
```

## Options

`-t, --title <title>`
: The title of the task.

`--listName [listName]`
: The name of the list in which to create the task. Specify either `listName` or `listId`, but not both.

`--listId [listId]`
: The id of the list in which to create the task. Specify either `listName` or `listId`, but not both.

`--bodyContent [bodyContent]`
: The body content of the task. In the UI this is called 'notes'.

`--bodyContentType [bodyContentType]`
: The type of the body content. Possible values are `text` and `html`. Default is `text`.

`--dueDateTime [dueDateTime]`
: The date when the task is due. This should be defined as a valid ISO 8601 string in the UTC time zone. Only date value is needed, time value is always ignored.

`--importance [importance]`
: The importance of the task. Possible values are: `low`, `normal`, `high`. Default is `normal`.

`--reminderDateTime [reminderDateTime]`
: The date and time for a reminder alert of the task to occur. This should be defined as a valid ISO 8601 string in the UTC time zone.

`--categories [categories]`
: Comma-separated list of categories associated with the task.

`--completedDateTime [completedDateTime]`
: The date and time when the task was finished. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`. This option can only be used when the `status` is set to `completed`.

`--startDateTime [startDateTime]`
: The date and time when the task is scheduled to start. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--status [status]`
: Indicates the state or progress of the task. The possible values are: `notStarted`, `inProgress`, `completed`, `waitingOnOthers`, `deferred`.

--8<-- "docs/cmd/_global.md"

## Remarks

When you specify the values for `categories`, each category can correspond to the displayName property of an [outlookCategory](https://learn.microsoft.com/graph/api/resources/outlookcategory?view=graph-rest-1.0). It is permissible to use distinct names.

## Examples

Add a task to Microsoft To Do tasks list with with a specific name

```sh
m365 todo task add --title "New task" --listName "My task list"
```

Add a task to a Microsoft To Do tasks list with a specific id

```sh
m365 todo task add --title "New task" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA=="
```

Create a new task with bodyContent and reminder and flag it as important

```sh
m365 todo task add --title "New task" --listName "My task list" --bodyContent "I should not forget this" --reminderDateTime 2023-01-01T12:00:00Z --importance high
```

Create a new task with a specific due date

```sh
m365 todo task add --title "New task" --listId "AQMkADlhMTRkOGEzLWQ1M2QtNGVkNS04NjdmLWU0NzJhMjZmZWNmMwAuAAADKvwNgAMNPE_zFNRJXVrU1wEAhHKQZHItDEOVCn8U3xuA2AABmQeVPwAAAA==" --dueDateTime 2023-01-01
```

Create a new task with categories, a completedDateTime, a startDateTime and a status

```sh
m365 todo task add --title "New task" --listName "My task list" --categories "Red category,Important" --completedDateTime 2023-12-01 --startDateTime 2023-12-01 --status "completed"
```

## Response

=== "JSON"

    ```json
    {
      "importance": "high",
      "isReminderOn": true,
      "status": "notStarted",
      "title": "New task",
      "createdDateTime": "2022-10-29T10:54:06.3672421Z",
      "lastModifiedDateTime": "2022-10-29T10:54:06.5078837Z",
      "hasAttachments": false,
      "categories": [],
      "id": "AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=",
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
    }
    ```

=== "Text"

    ```text
    body                : {"content":"I should not forget this","contentType":"text"}
    categories          : []
    createdDateTime     : 2022-10-29T10:54:06.3672421Z
    dueDateTime         : {"dateTime":"2023-01-01T00:00:00.0000000","timeZone":"UTC"}
    hasAttachments      : false
    id                  : AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=
    importance          : high
    isReminderOn        : true
    lastModifiedDateTime: 2022-10-29T10:54:06.5078837Z
    reminderDateTime    : {"dateTime":"2023-01-01T12:00:00.0000000","timeZone":"UTC"}
    status              : notStarted
    title               : New task
    ```

=== "CSV"

    ```csv
    importance,isReminderOn,status,title,createdDateTime,lastModifiedDateTime,hasAttachments,categories,id,body,dueDateTime,reminderDateTime
    high,1,notStarted,New task,2022-10-29T10:54:06.3672421Z,2022-10-29T10:54:06.5078837Z,,[],AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=,"{""content"":""I should not forget this"",""contentType"":""text""}","{""dateTime"":""2023-01-01T00:00:00.0000000"",""timeZone"":""UTC""}","{""dateTime"":""2023-01-01T12:00:00.0000000"",""timeZone"":""UTC""}"
    ```

=== "Markdown"

    ```md
    # todo task add --title "New task" --listName "My task list" --status "notStarted"

    Date: 4/3/2023

    ## New task (AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=)

    Property | Value
    ---------|-------
    importance | high
    isReminderOn | true
    status | notStarted
    title | New task
    createdDateTime | 2022-10-29T10:54:06.3672421Z
    lastModifiedDateTime | 2022-10-29T10:54:06.5078837Z
    hasAttachments | false
    categories | []
    id | AAMkAGYzNjMxYTU4LTJjZjYtNDlhMi1iMzQ2LWVmMTU3YmUzOGM5MABGAAAAAAAw3-tXgryDSr5p162KnUPKBwDEwEFouXWWT50CfwqSN9cpAAEX8ECDAADEwEFouXWWT50CfwqSN9cpAAEX8GuPAAA=
    body | {"content": "I should not forget this","contentType": "text"}
    dueDateTime | {"dateTime": "2023-01-01T00:00:00.0000000","timeZone": "UTC"}
    reminderDateTime |  {"dateTime": "2023-01-01T12:00:00.0000000","timeZone": "UTC"}
    ```
